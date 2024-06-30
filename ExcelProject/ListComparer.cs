namespace ExcelProject
{
    public class ListComparer : IEqualityComparer<List<string>>
    {
        public bool Equals(List<string> x, List<string> y)
        {
            if (x.Count != y.Count)
            {
                return false;
            }
            for (int i = 0; i < x.Count; i++)
            {
                if (x[i] != y[i])
                {
                    return false;
                }
            }
            return true;
        }

        public int GetHashCode(List<string> obj)
        {
            int hash = 17;
            foreach (var item in obj)
            {
                hash = hash * 23 + (item != null ? item.GetHashCode() : 0);
            }
            return hash;
        }
    }
}
