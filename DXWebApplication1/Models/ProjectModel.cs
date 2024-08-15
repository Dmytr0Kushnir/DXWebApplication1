using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;

namespace DXWebApplication1.Models
{
    public class ProjectModel
    {
        [BsonId]
        public ObjectId Id { get; set; }

        [BsonElement("project_name")]
        public string ProjectName { get; set; }

        [BsonElement("client_name")]
        public string ClientName { get; set; }

        [BsonElement("keystone_file_id")]
        public string KeyStoneFileId { get; set; }

        [BsonElement("project_address")]
        public string ProjectAddress { get; set; }

        [BsonElement("state_registration_id")]
        public string StateRegistrationId { get; set; }
    }
}