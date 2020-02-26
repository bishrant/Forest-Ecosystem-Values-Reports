using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace Report {
    // generated using this cool tool called app.quicktype.io
    public partial class ReportData {
        [JsonProperty("summaryReportResults")]
        public SummaryReportResults SummaryReportResults { get; set; }

        [JsonProperty("summaryResults")]
        public SummaryResult[] SummaryResults { get; set; }
    }

    public partial class SummaryReportResults {
        public object this[string propertyName] {
            get { return this.GetType().GetProperty(propertyName).GetValue(this, null); }
            set { this.GetType().GetProperty(propertyName).SetValue(this, value, null); }
        }

        [JsonProperty("airquality_forestAnnual")]
        public string airquality_forestAnnual { get; set; }

        [JsonProperty("airquality_forestValuePerAcre")]
        public string airquality_forestValuePerAcre { get; set; }

        [JsonProperty("airquality_ruralThousandDollarsPerYear")]
        public string airquality_ruralThousandDollarsPerYear { get; set; }

        [JsonProperty("airquality_totalThousandDollarsPerYear")]
        public string airquality_totalThousandDollarsPerYear { get; set; }

        [JsonProperty("airquality_urbanThousandDollarsPerYear")]
        public string airquality_urbanThousandDollarsPerYear { get; set; }

        [JsonProperty("aoiForestAcres")]
        public string AoiForestAcres { get; set; }

        [JsonProperty("aoiRuralAcres")]
        public string AoiRuralAcres { get; set; }

        [JsonProperty("aoiUrbanAcres")]
        public string AoiUrbanAcres { get; set; }

        [JsonProperty("biodiversity_forestAnnual")]
        public string biodiversity_forestAnnual { get; set; }

        [JsonProperty("biodiversity_forestValuePerAcre")]
        public string biodiversity_forestValuePerAcre { get; set; }

        [JsonProperty("biodiversity_ruralThousandDollarsPerYear")]
        public string biodiversity_ruralThousandDollarsPerYear { get; set; }

        [JsonProperty("biodiversity_totalThousandDollarsPerYear")]
        public string biodiversity_totalThousandDollarsPerYear { get; set; }

        [JsonProperty("biodiversity_urbanThousandDollarsPerYear")]
        public string biodiversity_urbanThousandDollarsPerYear { get; set; }

        [JsonProperty("carbon_forestAnnual")]
        public string carbon_forestAnnual { get; set; }

        [JsonProperty("carbon_forestValuePerAcre")]
        public string carbon_forestValuePerAcre { get; set; }

        [JsonProperty("carbon_ruralThousandDollarsPerYear")]
        public string carbon_ruralThousandDollarsPerYear { get; set; }

        [JsonProperty("carbon_totalThousandDollarsPerYear")]
        public string carbon_totalThousandDollarsPerYear { get; set; }

        [JsonProperty("carbon_urbanThousandDollarsPerYear")]
        public string carbon_urbanThousandDollarsPerYear { get; set; }

        [JsonProperty("cultural_forestAnnual")]
        public string cultural_forestAnnual { get; set; }

        [JsonProperty("cultural_forestValuePerAcre")]
        public string cultural_forestValuePerAcre { get; set; }

        [JsonProperty("cultural_ruralThousandDollarsPerYear")]
        public string cultural_ruralThousandDollarsPerYear { get; set; }

        [JsonProperty("cultural_totalThousandDollarsPerYear")]
        public string cultural_totalThousandDollarsPerYear { get; set; }

        [JsonProperty("cultural_urbanThousandDollarsPerYear")]
        public string cultural_urbanThousandDollarsPerYear { get; set; }

        [JsonProperty("total_forestAnnual")]
        public string total_forestAnnual { get; set; }

        [JsonProperty("total_forestValuePerAcre")]
        public string total_forestValuePerAcre { get; set; }

        [JsonProperty("total_ruralThousandDollarsPerYear")]
        public string total_ruralThousandDollarsPerYear { get; set; }

        [JsonProperty("total_totalThousandDollarsPerYear")]
        public string total_totalThousandDollarsPerYear { get; set; }

        [JsonProperty("total_urbanThousandDollarsPerYear")]
        public string total_urbanThousandDollarsPerYear { get; set; }

        [JsonProperty("watershed_forestAnnual")]
        public string watershed_forestAnnual { get; set; }

        [JsonProperty("watershed_forestValuePerAcre")]
        public string watershed_forestValuePerAcre { get; set; }

        [JsonProperty("watershed_ruralThousandDollarsPerYear")]
        public string watershed_ruralThousandDollarsPerYear { get; set; }

        [JsonProperty("watershed_totalThousandDollarsPerYear")]
        public string watershed_totalThousandDollarsPerYear { get; set; }

        [JsonProperty("watershed_urbanThousandDollarsPerYear")]
        public string watershed_urbanThousandDollarsPerYear { get; set; }
    }

    public partial class SummaryResult {
        [JsonProperty("AverageValueUnits")]
        public string AverageValueUnits { get; set; }

        [JsonProperty("TotalValueUnits")]
        public string TotalValueUnits { get; set; }

        [JsonProperty("ecosystemService")]
        public string EcosystemService { get; set; }

        [JsonProperty("forestAverageValue")]
        public string ForestAverageValue { get; set; }

        [JsonProperty("forestTotalValue")]
        public string ForestTotalValue { get; set; }

        [JsonProperty("ruralAverageValue")]
        public string RuralAverageValue { get; set; }

        [JsonProperty("ruralTotalValue")]
        public string RuralTotalValue { get; set; }

        [JsonProperty("urbanAverageValue")]
        public string UrbanAverageValue { get; set; }

        [JsonProperty("urbanTotalValue")]
        public string UrbanTotalValue { get; set; }
    }

    public partial class ReportData {
        public static ReportData FromJson(string json) => JsonConvert.DeserializeObject<ReportData>(json, Report.Converter.Settings);
    }

    public static class Serialize {
        public static string ToJson(this ReportData self) => JsonConvert.SerializeObject(self, Report.Converter.Settings);
    }

    internal static class Converter {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }

    internal class ParseStringConverter : JsonConverter {
        public override bool CanConvert(Type t) => t == typeof(long) || t == typeof(long?);

        public override object ReadJson(JsonReader reader, Type t, object existingValue, JsonSerializer serializer) {
            if (reader.TokenType == JsonToken.Null) return null;
            var value = serializer.Deserialize<string>(reader);
            long l;
            if (Int64.TryParse(value, out l)) {
                return l;
            }
            throw new Exception("Cannot unmarshal type long ");
        }

        public override void WriteJson(JsonWriter writer, object untypedValue, JsonSerializer serializer) {
            if (untypedValue == null) {
                serializer.Serialize(writer, null);
                return;
            }
            var value = (long)untypedValue;
            serializer.Serialize(writer, value.ToString());
            return;
        }

        public static readonly ParseStringConverter Singleton = new ParseStringConverter();
    }
}