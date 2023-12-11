using Google.Cloud.Firestore;
using System.Threading.Tasks;

namespace MyWebApi.Services
{

    public class FirestoreDbService : IFirestoreDb
    {
        private readonly IFirestoreDb _firestoreDb;

        public FirestoreDbService(IFirestoreDb firestoreDb)
        {
            _firestoreDb = firestoreDb ?? throw new ArgumentNullException(nameof(firestoreDb));
        }

        public CollectionReference Collection(string path)
        {
            //return Task.FromResult(_firestoreDb.Collection(path));
            return _firestoreDb.Collection(path);
        }
    }
}