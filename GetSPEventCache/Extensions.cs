using System.Data;
using System.Data.SqlClient;

namespace GetSPEventCache
{
    public static class Extension
    {
        public static bool CanOpen(SqlConnection connection)
        {
            try
            {
                if (connection == null) { return false; }

                connection.Open();
                var canOpen = connection.State == ConnectionState.Open;
                connection.Close();
                return canOpen;
            }
            catch
            {
                return false;
            }
        }
    }
}
