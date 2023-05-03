#if !NETPORTABLE_UIPATH
using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using SampleActivities.Basic.DataExtraction;
using SampleActivities.Basic.DocumentClassification;
using SampleActivities.Basic.OCR;

namespace SampleActivities
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();

            // Designers
            var simpleClassifierDesigner = new DesignerAttribute(typeof(SimpleClassifierDesigner));
//            var azureInvoiceDesigner = new DesignerAttribute(typeof(AzureInvoiceDesigner));
//            var clovaIDCardDesigner = new DesignerAttribute(typeof(ClovaIDCardDesigner));

            //Categories
            //var classifierCategoryAttribute = new CategoryAttribute("DU Extension Classifiers");
            var extractorCategoryAttribute = new CategoryAttribute("DU Extension Extractors");
            var ocrCategoryAttribute = new CategoryAttribute("DU Extension OCR Engines");

            //builder.AddCustomAttributes(typeof(SimpleClassifier), classifierCategoryAttribute);
            //builder.AddCustomAttributes(typeof(SimpleClassifier), simpleClassifierDesigner);

            builder.AddCustomAttributes(typeof(AzureInvoice), extractorCategoryAttribute);
//            builder.AddCustomAttributes(typeof(AzureInvoice), azureInvoiceDesigner);
            builder.AddCustomAttributes(typeof(ClovaDriverLicenseExtractor), extractorCategoryAttribute);
            builder.AddCustomAttributes(typeof(ClovaIDCardExtractor), extractorCategoryAttribute);
            builder.AddCustomAttributes(typeof(ClovaBusinessLicenseExtractor), extractorCategoryAttribute);

            builder.AddCustomAttributes(typeof(ClovaOCREngine), ocrCategoryAttribute);
            builder.AddCustomAttributes(typeof(HancomOCREngine), ocrCategoryAttribute);

            builder.AddCustomAttributes(typeof(ClovaOCREngine), nameof(ClovaOCREngine.Result), new CategoryAttribute("Output"));
            builder.AddCustomAttributes(typeof(HancomOCREngine), nameof(HancomOCREngine.Result), new CategoryAttribute("Output"));

            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
#endif