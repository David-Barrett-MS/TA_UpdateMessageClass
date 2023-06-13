/*
 *  Written by David Barrett, Microsoft Ltd. 2023.  Use at your own risk.  No warranties are given. 
 *  
 *  DISCLAIMER:
 * THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
 * MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
 * A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
 * MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
 * BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
 * SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
 * OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
 * */

using Microsoft.Exchange.Data.ContentTypes.Tnef;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Collections.Generic;

namespace TA_UpdateMessageClass
{
    /// <summary>
    /// Exchange 2019 routing agent showing how to update the message class of a message
    /// </summary>
    public class UpdateMessageClassAgent: RoutingAgent
    {
        static Logging.Logger _logger = null;
        Dictionary<string, string> _mapiProps = null;

        /// <summary>
        /// Constructor - instantiate our logger and subscribe to the OnSubmittedMessage event
        /// </summary>
        public UpdateMessageClassAgent()
        {
            // Create our logger (if we need to)
            if (_logger == null)
            {
                _logger = new Logging.Logger(false, $"c:\\TA\\UpdateMessageClassAgent.log");
                _logger.Log("UpdateMessageClassAgent instantiated");
            }

            base.OnSubmittedMessage += UpdateMessageClassAgent_OnSubmittedMessage;
        }

        /// <summary>
        /// Handle the OnSubmittedMessage event
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void UpdateMessageClassAgent_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs e)
        {
            _logger.Log($"OnSubmittedMessage called for item of type {e.MailItem.Message.MapiMessageClass}");
            if (e.MailItem.Message.Subject.StartsWith("UPDATEMESSAGECLASS") && e.MailItem.Message.MapiMessageClass == "IPM.Note")
            {
                UpdateMessageClass(e.MailItem);
                ParseMessageProps(e.MailItem);
                DumpAllProps();
            }
        }

        /// <summary>
        /// Apply message class IPM.Note.Custom to the message
        /// </summary>
        /// <param name="m">The MailItem to be updated</param>
        private void UpdateMessageClass(MailItem m)
        {
            _logger.Log("Attempting to update message class");

            try
            {
                Microsoft.Exchange.Data.Mime.MimePart tnefPart = m.Message.TnefPart;
                if (tnefPart != null)
                {
                    TnefReader reader = new TnefReader(tnefPart.GetContentReadStream());
                    TnefWriter writer = new TnefWriter(
                        tnefPart.GetContentWriteStream(tnefPart.ContentTransferEncoding),
                        reader.AttachmentKey);
                    while (reader.ReadNextAttribute())
                    {
                        if (reader.AttributeTag == TnefAttributeTag.MapiProperties)
                        {
                            writer.StartAttribute(TnefAttributeTag.MapiProperties, TnefAttributeLevel.Message);
                            while (reader.PropertyReader.ReadNextProperty())
                            {
                                if (!reader.PropertyReader.IsNamedProperty)
                                {
                                    // Check if this is message class
                                    if (reader.PropertyReader.PropertyTag.ToString().Equals("MessageClass (Unicode)"))
                                    {
                                        _logger.Log($"Setting {reader.PropertyReader.PropertyTag} to IPM.Note.Custom");
                                        writer.StartProperty(reader.PropertyReader.PropertyTag);
                                        writer.WritePropertyValue("IPM.Note.Custom");
                                    }
                                    else
                                    {
                                        writer.WriteProperty(reader.PropertyReader);
                                    }
                                }
                                else
                                    writer.WriteProperty(reader.PropertyReader);
                            }
                        }
                        else
                        {
                            writer.WriteAttribute(reader);
                        }
                    }
                    reader.Close();
                    writer.Close();
                }
            }
            catch (Exception ex)
            {
                _logger.Log($"Error updating TNEF: {ex.Message}");
            }

        }

        /// <summary>
        /// Dump all TNEF properties to the log file
        /// </summary>
        private void DumpAllProps()
        {
            // Write all properties to the log file
            if (_mapiProps == null)
            {
                _logger.Log("No MAPI properties found on message");
                return;
            }

            _logger.Log("Message properties:");
            foreach (string propName in _mapiProps.Keys)
            {
                try
                {
                    _logger.Log($"{propName} = {_mapiProps[propName]}");
                }
                catch { }
            }
            _logger.Log("Prop dump complete");
        }

        /// <summary>
        /// Parse the TNEF message properties and dump them into a Dictionary
        /// </summary>
        /// <param name="mailItem">The MailItem to process</param>
        private void ParseMessageProps(MailItem mailItem)
        {
            if (mailItem == null)
            {
                _logger.Log("Null message passed to ParseMessageProps");
                return;
            }

            Microsoft.Exchange.Data.Mime.MimePart TNEFMIMEPart = mailItem.Message.TnefPart;
            _mapiProps = new Dictionary<string, string>();
            if (TNEFMIMEPart != null)
            {
                // Use the TNEF reader to parse the message properties
                try
                {
                    TnefReader tnefReader = new TnefReader(TNEFMIMEPart.GetContentReadStream(), 1252, TnefComplianceMode.Loose);
                    _logger.Log("TNEF stream reader created");
                    while (tnefReader.ReadNextAttribute())
                    {
                        try
                        {
                            if (tnefReader.AttributeTag == TnefAttributeTag.MapiProperties)
                            {
                                _logger.Log("MapiProperties located");
                                while (tnefReader.PropertyReader.ReadNextProperty())
                                {
                                    string propTagOrId = String.Empty;
                                    string propValue = String.Empty;
                                    try
                                    {
                                        if (tnefReader.PropertyReader.IsNamedProperty)
                                        {
                                            propTagOrId = tnefReader.PropertyReader.PropertyNameId.ToString();
                                        }
                                        else
                                        {
                                            propTagOrId = "MAPI:" + tnefReader.PropertyReader.PropertyTag.ToString();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        _logger.Log("Error while reading property Id: " + ex.Message);
                                    }
                                    try
                                    {
                                        propValue = tnefReader.PropertyReader.ReadValue().ToString();
                                    }
                                    catch (Exception ex)
                                    {
                                        _logger.Log("Error while reading property value: " + ex.Message);
                                    }

                                    if (!String.IsNullOrEmpty(propTagOrId))
                                    {
                                        if (String.IsNullOrEmpty(propValue))
                                            propValue = "unknown";
                                        if (!_mapiProps.ContainsKey(propTagOrId))
                                            _mapiProps.Add(propTagOrId, propValue);
                                    }
                                }
                                _logger.Log($"MapiProperties read: {_mapiProps.Count}");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Log("Error while parsing attributes: " + ex.Message);
                        }
                    }
                    tnefReader.Close();
                }
                catch (Exception ex)
                {
                    _logger.Log("Error: " + ex.Message);
                }
            }
        }

    }
}
