using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RemotePrint.Controllers
{
    public class PlcEnumerator
    {
        private readonly Stream _stream;
        private readonly byte[] _buffer = new byte[5242880];
        private readonly List<byte> _page = new List<byte>();
        private const int ScFf = 0x000C; // Form-feed

        public PlcEnumerator(Stream stream)
        {
            _stream = stream;
        }

        public bool GetPage(out byte[] pageBytes)
        {
            //Create a dummy anonymous variable 'Where(arg => 1 == 2)' makes sure we don't get any rows and its null
            var nextFormFeed =
                _page.Select((v, i) => new {b = v, Index = i}).Where(arg => 1 == 2).FirstOrDefault(b => b.b == ScFf);

            while (nextFormFeed == null)
            {
                //Do we already have a form feed
                nextFormFeed = _page.Select((v, i) => new {b = v, Index = i}).FirstOrDefault(b => b.b == ScFf);

                if (nextFormFeed != null)
                {
                    //Print the page
                    var indexOfFormFeed = nextFormFeed.Index;

                    var currentPageBytes = _page.GetRange(0, indexOfFormFeed + 1).ToArray(); //+ 1 to make sure to include the Form Feed

                    //Remove printed data
                    _page.RemoveRange(0, indexOfFormFeed + 1); //+ 1 to make sure to remove previous Form Feed
                    pageBytes = currentPageBytes;
                    return true;
                }

                //Didnt find a form feed in buffer so far, read more data
                if (_stream.Read(_buffer, 0, _buffer.Length) > 0)
                {
                    _page.AddRange(_buffer);
                }
                else
                {
                    //Nothing more to read
                    pageBytes = null;
                    return false;
                }
            }

            //Not sure if we will ever get to this code, may be when its a empty file
            pageBytes = null;
            return false;
        }
    }
}