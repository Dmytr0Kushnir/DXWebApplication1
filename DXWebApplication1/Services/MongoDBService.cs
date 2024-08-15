using MongoDB.Driver;
using MongoDB.Bson;
using System.Collections.Generic;
using DXWebApplication1.Models;

public class MongoDBService
{
    private readonly IMongoCollection<ProjectModel> _projectCollection;
    private readonly IMongoCollection<CompModel> _compsCollection;

    public MongoDBService()
    {
        string connectionString = "mongodb+srv://illia_pv:localhost1617@cluster0.nvlushg.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
        string databaseName = "keystone_db";
        string projectsCollectionName = "projects";
        string compsCollectionName = "comp";

        var client = new MongoClient(connectionString);
        var database = client.GetDatabase(databaseName);
        _projectCollection = database.GetCollection<ProjectModel>(projectsCollectionName);
        _compsCollection = database.GetCollection<CompModel>(compsCollectionName);

    }

    public List<CompModel> GetComps()
    {
        return _compsCollection.Find(comp => true).ToList();
    }

    public List<ProjectModel> GetProjects()
    {
        return _projectCollection.Find(project => true).ToList();
    }

}
