using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportLobo
{
    class Program
    {
        static void Main(string[] args)
        {
            // DOC_DOCUMENTS --> Displayname + ID
            // Folder = COMPANY

            try
            {
                DMS1.Instance inst = new DMS1.Instance();
                inst.Connect(Properties.Settings.Default.LoboConnectionString);

                DMS1.DocumentQuery qry = inst.Category("DOC_DOCUMENTS").NewDocumentQuery();
                DMS1.DocumentSet ds = qry.AllDocuments();

                int cur = 0;
                int max = ds.Count();

                while(!ds.EOF)
                {
                    try
                    {
                        DMS1.Document doc = ds.Current();

                        string name = doc.Name();
                        string company = Convert.ToString(doc.IndexDataReadonly.DisplayValue["COMPANY", true]);
                        string expName = string.Format("{0} {1}.{2}", InvalidFile(name), doc.ID, doc.Content().FileExtension().TrimStart('.'));

                        string expFolder = InvalidPath(System.IO.Path.Combine(Properties.Settings.Default.ExportFolder, company));
                        string destFile = GetUniqueFilePaht(System.IO.Path.Combine(expFolder, expName));

                        System.IO.Directory.CreateDirectory(expFolder);

                        Console.Write(string.Format("Exporting: '{0}' to '{1}'", name, destFile));
                        string retFile = doc.Content().Retrieve();
                        System.IO.File.Move(retFile, destFile);
                        Console.WriteLine(" - DONE");
                    }
                    finally
                    {
                        Console.Title = string.Format("{0}/{1} processed", ++cur, max);
                        ds.MoveNext();
                    }

                    
                }
            }
            catch(Exception globEx)
            {
                Console.WriteLine(string.Format("GLOBAL ERROR: {0}{1}{2}", globEx.Message, Environment.NewLine, globEx.ToString()));
            }
        }

        public static string InvalidPath(string path)
        {
            string ret = path;
            foreach (char entry in System.IO.Path.GetInvalidPathChars())
                ret = ret.Replace(entry, '_');

            if (ret.Length > 75)
                ret = ret.Substring(0, 75);

            return ret;
        }

        public static string InvalidFile(string file)
        {
            string ret = file;
            foreach (char entry in System.IO.Path.GetInvalidFileNameChars())
                ret = ret.Replace(entry, '_');

            if (ret.Length > 75)
                ret = ret.Substring(0, 75);

            return ret;
        }

        public static string GetUniqueFilePaht(string file)
        {
            if (!System.IO.File.Exists(file))
                return file;

            int curCount = 1;
            string fileName = System.IO.Path.GetFileName(file);
            string ext = System.IO.Path.GetExtension(file).TrimStart('.');

            string newFile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(file), string.Format("{0} ({1}).{2}", fileName, curCount++, ext));

            while(System.IO.File.Exists(newFile))
                newFile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(file), string.Format("{0} ({1}).{2}", fileName, curCount++, ext));

            return newFile;
        }
    }
}
