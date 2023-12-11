using Google.Cloud.Firestore;

public interface ICollectionReference
{
    Task<DocumentReference> Add(object data);
    DocumentReference Document();
    DocumentReference Document(string documentPath);
    string GetId();
    DocumentReference GetParent();
    string GetPath();
}
