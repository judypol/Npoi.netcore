﻿/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System.IO;
using npoi.core.Corefx;
using NPOI.POIFS.FileSystem;
using Xunit;

namespace npoi.test.poifs
{
    public class ReadThumbsDB
    {
        [Fact]
        public void Main()
        {
            FileStream stream = new FileStream("thumbs.db", FileMode.Open, FileAccess.Read);
            POIFSFileSystem poifs = new POIFSFileSystem(stream);
            var entries = poifs.Root.Entries;
            //POIFSDocumentReader catalogdr = poifs.CreatePOIFSDocumentReader("Catalog");
            //byte[] b1=new byte[catalogdr.Length-4];
            //catalogdr.Read(b1,4,b1.Length);
            //Dictionary<string, string> indexList = new Dictionary<string, string>();
            //for (int j = 0; j < b1.Length; j++)
            //{ 
            //    if(b1[0]
            //}

            while (entries.MoveNext())
            {
                DocumentNode entry = entries.Current as DocumentNode;
                DocumentInputStream dr = poifs.CreateDocumentInputStream(entry.Name);

                if (entry.Name.ToLower() == "catalog")
                    continue;

                byte[] buffer = new byte[dr.Length];
                dr.Read(buffer);
                int startpos = 0;

                //detect jfif header
                for (int i = 3; i < buffer.Length; i++)
                {
                    if (buffer[i - 3] == 0xFF
                        && buffer[i - 2] == 0xD8
                        && buffer[i - 1] == 0xFF
                        && buffer[i] == 0xE0)
                    {
                        startpos = i - 3;
                        break;
                    }
                }
                if (startpos == 0)
                    continue;

                FileStream jpeg = File.Create(entry.Name + ".jpeg");
                jpeg.Write(buffer, startpos, buffer.Length - startpos);
                jpeg.Close();
            }

            stream.Close();
        }
    }
}
