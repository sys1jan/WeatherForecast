using Google.Cloud.Firestore;

public interface IFirestoreDb
{
    CollectionReference Collection(string path);
}
