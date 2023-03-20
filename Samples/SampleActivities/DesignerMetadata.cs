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
            var azureInvoiceDesigner = new DesignerAttribute(typeof(AzureInvoiceDesigner));

            //Categories
            var classifierCategoryAttribute = new CategoryAttribute("Sample Classifiers");
            var extractorCategoryAttribute = new CategoryAttribute("Azure Extractors");
            var ocrCategoryAttribute = new CategoryAttribute("Sample OCR Engines");

            builder.AddCustomAttributes(typeof(SimpleClassifier), classifierCategoryAttribute);
            builder.AddCustomAttributes(typeof(SimpleClassifier), simpleClassifierDesigner);

            builder.AddCustomAttributes(typeof(AzureInvoice), extractorCategoryAttribute);
            builder.AddCustomAttributes(typeof(AzureInvoice), azureInvoiceDesigner);

            builder.AddCustomAttributes(typeof(SimpleOCREngine), ocrCategoryAttribute);
            builder.AddCustomAttributes(typeof(SimpleOCREngine), nameof(SimpleOCREngine.Result), new CategoryAttribute("Output"));

            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
#endif