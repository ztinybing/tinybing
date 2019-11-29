using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Com.Bing
{
    class ListItem
    {
        private string fullPath;
        public string FullPath
        {
            get
            {
                return this.fullPath;
            }
        }
        public string ParentDirPath
        {
            get
            {
                return Path.GetDirectoryName(this.fullPath);
            }
        }
        public string NewFolder
        {
            get
            {
                string newFolder = string.Concat(this.ParentDirPath, string.Concat("\\", this.FileName));
                if (!Directory.Exists(newFolder))
                {
                    Directory.CreateDirectory(newFolder);
                }
                return newFolder;
            }
        }
        public string FileName
        {
            get
            {
                return Path.GetFileNameWithoutExtension(this.fullPath);
            }
        }
        public ListItem(string fullPath)
        {
            this.fullPath = fullPath;
        }
        public override string ToString()
        {
            return this.FileName;
        }
    }
}
