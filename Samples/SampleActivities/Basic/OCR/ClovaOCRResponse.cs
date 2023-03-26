using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;

namespace SampleActivities.Basic.OCR
{
    enum SynapFieldType
    {
        DEFAULT = 1,
        BLOCK = 2,
        LINE = 4,
        MATCHED = 8
    }


    public class SynapOCRField
    {
        public List<PointF> Points { get; } = new List<PointF>();

        public double Confidence { get; set; }
        public string Text { get; set; }

        public override string ToString()
        {
            return string.Join(", ", Points) + $", confidence: {Confidence}, Text: {Text}";
        }
    }

    public class SynapFormField
    {
        public string Id { get; set;  }
        public string Label { get; set; }

        public string Value { get; set; }
        //public List<PointF> Points { get; } = new List<PointF>();
        public string Masking { get; set; }
        public double Confidence { get; set; }

        public override string ToString()
        {
            return $", fid: {Id}, label: {Label}, value : {Value}, masking: {Masking}, confidence: {Confidence}";
        }
    }
    public class SynapForm
    {
        public string id { get; set; }
        public double confidence { get; set; }
        public string name { get; set; }
        public string type { get; set;  }
        //public List< List<PointF>> boundary { get; } = new List< List<PointF>>();
        public List<object> form_list { get; set; }
        public SynapFormField[] fields
        {
            get
            {
                return GetFields();
            }
        }

        private List<SynapFormField> _form_list { get; set; }

        public override string ToString()
        {
            return $", id: {id}, confidene: {confidence}, name: {name}, type: {type}";
        }
        private void parseFormField(List<SynapFormField> fields, List<object> items)
        {
            foreach (var obj in items)
            {
                var field = new SynapFormField();
                JArray arr = JArray.Parse(obj.ToString());
                field.Id = arr[0].ToString();
                field.Label = arr[1].ToString();
                field.Value = arr[2].ToString();
                /*
                for (int i = 0; i < 4; i++)
                {
                    var tmp = JArray.Parse(arr[i + 3].ToString());
                    field.Points.Add(new PointF(Single.Parse(tmp[0].ToString()), Single.Parse(tmp[1].ToString())));
                }
                */
                field.Masking = arr[4].ToString();
                field.Confidence = Single.Parse(arr[5].ToString());
                fields.Add(field);
            }
        }
        public SynapFormField[] GetFields()
        {
            if (_form_list != null)
                return _form_list.ToArray();
            _form_list = new List<SynapFormField>();
            parseFormField( _form_list, form_list);
            return _form_list.ToArray();
        }
    }

    public class SynapOCRResult
    {
        public string fid { get; set; }
        public double dur { get; set; }
        public string csv_file_name { get; set; }
        public string full_text { get; set; }
        public int height { get; set; }
        public int width { get; set; }
        public string final_file_name { get; set; }
        public int total_page { get; set; }
        public int page_index { get; set; }
        public double rotate { get; set; }
        public string masked_image { get; set; }

        private void parseFields(List<SynapOCRField> fields, List<object> items)
        {
            foreach (var obj in items)
            {
                var field = new SynapOCRField();
                JArray arr = JArray.Parse(obj.ToString());
                for (int i = 0; i < 4; i++)
                {
                    var tmp = JArray.Parse(arr[i].ToString());
                    field.Points.Add(new PointF(Single.Parse(tmp[0].ToString()), Single.Parse(tmp[1].ToString())));
                }
                field.Confidence = Convert.ToInt32((100*Double.Parse(arr[4].ToString())));
                field.Text = arr[5].ToString();
                fields.Add(field);
            }
        }


        private void parseForms(  List<object> items)
        {
            if (this._matched_forms == null)
                this._matched_forms = new List<SynapForm>();
            if (this._matched_forms.Count() > 0)
                this._matched_forms.Clear();
            foreach (var inner in items)
            {
                JArray ineers = JArray.Parse(inner.ToString());
                foreach (var obj in ineers)
                {
                    SynapForm form = JsonConvert.DeserializeObject<SynapForm>(obj.ToString(), new SynapFormConverter());
                this._matched_forms.Add(form);
                }
            }
        }
        public SynapOCRField[] GetFields(int type)
        {
            if (type == (int)SynapFieldType.DEFAULT)
            {
                if (this._fields != null)
                    return this._fields.ToArray();
                this._fields = new List<SynapOCRField>();
                parseFields(this._fields, this.boxes);
                return this._fields.ToArray();
            }
            else if (type == (int)SynapFieldType.BLOCK)
            {
                if (this._block_fields != null)
                    return this._block_fields.ToArray();
                this._block_fields = new List<SynapOCRField>();
                parseFields(this._block_fields, this.block_boxes);
                return this._block_fields.ToArray();
            }
            else if (type == (int)SynapFieldType.LINE)
            {
                if (this._line_fields != null)
                    return this._line_fields.ToArray();
                this._line_fields = new List<SynapOCRField>();
                parseFields(this._line_fields, this.line_boxes);
                return this._line_fields.ToArray();
            }
            else if( type == (int) SynapFieldType.MATCHED)
            {
                if( this._matched_fields != null)
                    return this._matched_fields.ToArray();
                this._matched_fields = new List<SynapOCRField>();
                parseFields( this._matched_fields, this.matched_boxes);
                return this._matched_fields.ToArray();

            }
            return null;
        }

        public SynapForm[] GetSynapForms()
        {
            parseForms(this.matched_forms);
            return this._matched_forms.ToArray();
        }


        private List<SynapOCRField> _fields;
        private List<SynapOCRField> _block_fields;
        private List<SynapOCRField> _line_fields;
        private List<SynapOCRField> _matched_fields;

        private List<SynapForm> _matched_forms;

        public List<object> boxes { get; set; }
        public List<object> block_boxes { get; set; }
        public List<object> line_boxes { get; set; }
        public List<object> matched_boxes { get; set;}
        public List<object> matched_forms { get; set;}

        public override string ToString()
        {
            return $"FID: {fid}, full_text: {full_text},  Dur: {dur}, height: {height}, width: {width}, total_page: {total_page}, page_index: {page_index}, final_file_name: {final_file_name}, csv_file_name: {csv_file_name} rotate: {rotate} boxes: {boxes} ";
        }

    }
    public class ClovaOCRResponse
    {
        public int status { get; set; }
        public SynapOCRResult result { get; set; }

        public override string ToString()
        {
            return $"status: {status}, result : {result}";
        }
    }

    public class SynapOCRResponseConverter : CustomCreationConverter<ClovaOCRResponse>
    {
        public override ClovaOCRResponse Create(Type objectType)
        {
            return new ClovaOCRResponse();
        }
    }

    public class SynapFormConverter : CustomCreationConverter<SynapForm>
    {
        public override SynapForm Create(Type objectType)
        {
            return new SynapForm();
        }
    }
}
