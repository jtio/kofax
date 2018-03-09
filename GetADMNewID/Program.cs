using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.InteropServices;


namespace GetADMNewID
{
    //[ComVisible(true),
    //GuidAttribute("D9675881-BC69-4e52-A45D-08F1025FFC28")]
    //[InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(KTMUtils.InterfaceId)]
    public interface IKTMUtils
    {
        string GetNewID(string url);
        void Connect(string url);
        
    }

    [Guid(KTMUtils.EventsId), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IKTMEvents { }

    /// <summary>
    /// Class that implements various
    /// </summary>
    //[ComVisible(true),
    //Guid("85EA27F3-0254-438A-81A9-599E10F4005C"),
    //    ClassInterface(ClassInterfaceType.None)]
    //[ProgId("ImageSource.KTMUtils")]
    [Guid(KTMUtils.ClassId), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(IKTMEvents))]
    public class KTMUtils : IKTMUtils
    {
        string WebServicesURL;

        #region declarations and process entry point
        public const string ClassId = "EFFC0B5C-F240-4361-9D1C-F8FB7311EE08";
        public const string InterfaceId = "18098E99-C380-4509-AAC0-0EA96848E8BF";
        public const string EventsId = "2C278B78-00AE-4DD1-AE11-9A4035344BE8";

        public KTMUtils() { }

        #endregion
        #region execute methods
        /// <summary>
        /// Retrieve new document ID from Advance ADM.
        /// </summary>
        public string GetNewID(string url)
        {
            edu.gatech.gtf.aridva.ADMIntegration NewID = new edu.gatech.gtf.aridva.ADMIntegration(url);
            edu.gatech.gtf.aridva.RequestNewDocumentIdResult resp;


            resp = NewID.RequestNewDocumentId();
            string newid = resp.NewDocumentId;
            
            return newid;
        }

        public void Connect(string url)
        {
            WebServicesURL = url;
        }

        #endregion
    }
}
