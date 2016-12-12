// ------------------------------------------------------------------------------
//  <auto-generated>
//    Generated by Xsd2Code. Version 3.4.0.38967
//    <NameSpace>schemas</NameSpace><Collection>List</Collection><codeType>CSharp</codeType><EnableDataBinding>False</EnableDataBinding><EnableLazyLoading>False</EnableLazyLoading><TrackingChangesEnable>False</TrackingChangesEnable><GenTrackingClasses>False</GenTrackingClasses><HidePrivateFieldInIDE>False</HidePrivateFieldInIDE><EnableSummaryComment>False</EnableSummaryComment><VirtualProp>False</VirtualProp><IncludeSerializeMethod>False</IncludeSerializeMethod><UseBaseClass>False</UseBaseClass><GenBaseClass>False</GenBaseClass><GenerateCloneMethod>False</GenerateCloneMethod><GenerateDataContracts>False</GenerateDataContracts><CodeBaseTag>Net20</CodeBaseTag><SerializeMethodName>Serialize</SerializeMethodName><DeserializeMethodName>Deserialize</DeserializeMethodName><SaveToFileMethodName>SaveToFile</SaveToFileMethodName><LoadFromFileMethodName>LoadFromFile</LoadFromFileMethodName><GenerateXMLAttributes>False</GenerateXMLAttributes><EnableEncoding>False</EnableEncoding><AutomaticProperties>False</AutomaticProperties><GenerateShouldSerialize>False</GenerateShouldSerialize><DisableDebug>False</DisableDebug><PropNameSpecified>Default</PropNameSpecified><Encoder>UTF8</Encoder><CustomUsings></CustomUsings><ExcludeIncludedTypes>False</ExcludeIncludedTypes><EnableInitializeFields>True</EnableInitializeFields>
//  </auto-generated>
// ------------------------------------------------------------------------------
namespace NPOI.OpenXmlFormats
{
    using NPOI.OpenXml4Net.Util;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Xml;
    using System.Xml.Serialization;


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/schemaLibrary/2006/main")]
    public class CT_SchemaLibrary
    {

        private List<CT_Schema> schemaField;

        public CT_SchemaLibrary()
        {
            this.schemaField = new List<CT_Schema>();
        }

        [XmlElement("schema")]
        public List<CT_Schema> schema
        {
            get
            {
                return this.schemaField;
            }
            set
            {
                this.schemaField = value;
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/schemaLibrary/2006/main")]
    public partial class CT_Schema
    {

        private string uriField;

        private string manifestLocationField;

        private string schemaLocationField;
        public static CT_Schema Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Schema ctObj = new CT_Schema();
            ctObj.uri = XmlHelper.ReadString(node.Attributes["w:uri"]);
            ctObj.manifestLocation = XmlHelper.ReadString(node.Attributes["w:manifestLocation"]);
            ctObj.schemaLocation = XmlHelper.ReadString(node.Attributes["w:schemaLocation"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:uri", this.uri);
            XmlHelper.WriteAttribute(sw, "w:manifestLocation", this.manifestLocation);
            XmlHelper.WriteAttribute(sw, "w:schemaLocation", this.schemaLocation);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_Schema()
        {
            this.uriField = "";
        }

        [DefaultValueAttribute("")]
        [XmlAttribute]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }

        [XmlAttribute]
        public string manifestLocation
        {
            get
            {
                return this.manifestLocationField;
            }
            set
            {
                this.manifestLocationField = value;
            }
        }

        [XmlAttribute]
        public string schemaLocation
        {
            get
            {
                return this.schemaLocationField;
            }
            set
            {
                this.schemaLocationField = value;
            }
        }
    }
}
