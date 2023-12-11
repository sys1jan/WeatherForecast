using Google.Cloud.Firestore;

public class CollectionReferenceWrapper : ICollectionReference
{
    private readonly CollectionReference _collectionReference;

    public CollectionReferenceWrapper(CollectionReference collectionReference)
    {
        _collectionReference = collectionReference;
    }

    public Task<DocumentReference> Add(object data)
    {
        return _collectionReference.AddAsync(data);
    }

    public DocumentReference Document()
    {
        return _collectionReference.Document();
    }

    public DocumentReference Document(string documentPath)
    {
        return _collectionReference.Document(documentPath);
    }

    public string GetId()
    {
        return _collectionReference.Id;
    }

    public DocumentReference GetParent()
    {
        return _collectionReference.Parent;
    }

    public string GetPath()
    {
        return _collectionReference.Path;
    }
}