namespace RangerComBrowser
{
    public struct WebItem
    {
        public int Index { get; set; }
        public string Value { get; set; }

        public WebItem(int index, string value)
        {
            this.Index = index;
            this.Value = value;
        }
    }
}
