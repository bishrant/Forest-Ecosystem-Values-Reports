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
        [JsonProperty("airquality_forestAnnual")]
        [JsonConverter(typeof(ParseStringConverter))]
        public long AirqualityForestAnnual { get; set; }

        [JsonProperty("airquality_forestValuePerAcre")]
        public string AirqualityForestValuePerAcre { get; set; }

        [JsonProperty("airquality_ruralThousandDollarsPerYear")]
        public string AirqualityRuralThousandDollarsPerYear { get; set; }

        [JsonProperty("airquality_totalThousandDollarsPerYear")]
        public string AirqualityTotalThousandDollarsPerYear { get; set; }

        [JsonProperty("airquality_urbanThousandDollarsPerYear")]
        public string AirqualityUrbanThousandDollarsPerYear { get; set; }

        [JsonProperty("aoiForestAcres")]
        public string AoiForestAcres { get; set; }

        [JsonProperty("aoiRuralAcres")]
        public string AoiRuralAcres { get; set; }

        [JsonProperty("aoiUrbanAcres")]
        [JsonConverter(typeof(ParseStringConverter))]
        public long AoiUrbanAcres { get; set; }

        [JsonProperty("biodiversity_forestAnnual")]
        public string BiodiversityForestAnnual { get; set; }

        [JsonProperty("biodiversity_forestValuePerAcre")]
        public string BiodiversityForestValuePerAcre { get; set; }

        [JsonProperty("biodiversity_ruralThousandDollarsPerYear")]
        public string BiodiversityRuralThousandDollarsPerYear { get; set; }

        [JsonProperty("biodiversity_totalThousandDollarsPerYear")]
        public string BiodiversityTotalThousandDollarsPerYear { get; set; }

        [JsonProperty("biodiversity_urbanThousandDollarsPerYear")]
        public string BiodiversityUrbanThousandDollarsPerYear { get; set; }

        [JsonProperty("carbon_forestAnnual")]
        public string CarbonForestAnnual { get; set; }

        [JsonProperty("carbon_forestValuePerAcre")]
        public string CarbonForestValuePerAcre { get; set; }

        [JsonProperty("carbon_ruralThousandDollarsPerYear")]
        public string CarbonRuralThousandDollarsPerYear { get; set; }

        [JsonProperty("carbon_totalThousandDollarsPerYear")]
        public string CarbonTotalThousandDollarsPerYear { get; set; }

        [JsonProperty("carbon_urbanThousandDollarsPerYear")]
        public string CarbonUrbanThousandDollarsPerYear { get; set; }

        [JsonProperty("cultural_forestAnnual")]
        public string CulturalForestAnnual { get; set; }

        [JsonProperty("cultural_forestValuePerAcre")]
        public string CulturalForestValuePerAcre { get; set; }

        [JsonProperty("cultural_ruralThousandDollarsPerYear")]
        public string CulturalRuralThousandDollarsPerYear { get; set; }

        [JsonProperty("cultural_totalThousandDollarsPerYear")]
        public string CulturalTotalThousandDollarsPerYear { get; set; }

        [JsonProperty("cultural_urbanThousandDollarsPerYear")]
        public string CulturalUrbanThousandDollarsPerYear { get; set; }

        [JsonProperty("total_forestAnnual")]
        public string TotalForestAnnual { get; set; }

        [JsonProperty("total_forestValuePerAcre")]
        public string TotalForestValuePerAcre { get; set; }

        [JsonProperty("total_ruralThousandDollarsPerYear")]
        public string TotalRuralThousandDollarsPerYear { get; set; }

        [JsonProperty("total_totalThousandDollarsPerYear")]
        public string TotalTotalThousandDollarsPerYear { get; set; }

        [JsonProperty("total_urbanThousandDollarsPerYear")]
        public string TotalUrbanThousandDollarsPerYear { get; set; }

        [JsonProperty("watershed_forestAnnual")]
        public string WatershedForestAnnual { get; set; }

        [JsonProperty("watershed_forestValuePerAcre")]
        public string WatershedForestValuePerAcre { get; set; }

        [JsonProperty("watershed_ruralThousandDollarsPerYear")]
        public string WatershedRuralThousandDollarsPerYear { get; set; }

        [JsonProperty("watershed_totalThousandDollarsPerYear")]
        public string WatershedTotalThousandDollarsPerYear { get; set; }

        [JsonProperty("watershed_urbanThousandDollarsPerYear")]
        public string WatershedUrbanThousandDollarsPerYear { get; set; }
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
            throw new Exception("Cannot unmarshal type long");
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