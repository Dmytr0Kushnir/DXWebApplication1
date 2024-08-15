using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;

namespace DXWebApplication1.Models
{
    public class CompModel
    {
        [BsonId]
        public ObjectId Id { get; set; }

        [BsonElement("sale_number")]
        public string SaleNumber { get; set; }

        [BsonElement("property_type")]
        public string PropertyType { get; set; }

        [BsonElement("address")]
        public string Address { get; set; }

        [BsonElement("city_state")]
        public string CityState { get; set; }

        [BsonElement("verification")]
        public string Verification { get; set; }

        [BsonElement("sale_price")]
        public string SalePrice { get; set; }

        [BsonElement("sale_date")]
        public string SaleDate { get; set; }

        [BsonElement("sq_ft")]
        public string SqFt { get; set; }

        [BsonElement("price_per_sq_ft")]
        public string PricePerSqFt { get; set; }

        [BsonElement("site_size_acres")]
        public string SiteSizeAcres { get; set; }

        [BsonElement("grantor")]
        public string Grantor { get; set; }

        [BsonElement("location")]
        public string Location { get; set; }

        [BsonElement("year_built")]
        public string YearBuilt { get; set; }

        [BsonElement("condition")]
        public string Condition { get; set; }

        [BsonElement("access_visibility")]
        public string AccessVisibility { get; set; }

        [BsonElement("tenancy")]
        public string Tenancy { get; set; }

        [BsonElement("land_to_building_ratio")]
        public string LandToBuildingRatio { get; set; }

        [BsonElement("cap_rate")]
        public string CapRate { get; set; }

        [BsonElement("interior_finish")]
        public string InteriorFinish { get; set; }

        [BsonElement("quality_of_construction")]
        public string QualityOfConstruction { get; set; }

        [BsonElement("grantee")]
        public string Grantee { get; set; }

        [BsonElement("project_id")]
        [BsonRepresentation(BsonType.ObjectId)]
        public string ProjectId { get; set; }

    }
}
