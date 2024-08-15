using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DXWebApplication1.Models
{
    public class ProjectOptionModel
    {
        [BsonId]
        public ObjectId Id { get; set; }

        [BsonElement("project_name")]
        public string ProjectName { get; set; }
    }
}