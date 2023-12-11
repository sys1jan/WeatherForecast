using Google.Cloud.Firestore;

public class FirestoreDbWrapper : IFirestoreDb
{
    private readonly IFirestoreDb _firestoreDb;

    public FirestoreDbWrapper(IFirestoreDb firestoreDb)
    {
        _firestoreDb = firestoreDb;
    }

    public CollectionReference Collection(string path)
    {
        return _firestoreDb.Collection(path);
    }
}
