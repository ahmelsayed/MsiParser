using Microsoft.Deployment.WindowsInstaller;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MsiParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var db = Parse(args[0]);
            db.Files
                .Select(f => GetFullPath(db, db.Components.Where(c => c.Component == f.Component).Select(c => c.Directory).FirstOrDefault(), f.FileName.Split('|').Last()))
                .ToList()
                .ForEach(Console.WriteLine);
        }

        static MsiDatabase Parse(string path)
        {
            using (var database = new Database(path, DatabaseOpenMode.ReadOnly))
            {
                var db = new MsiDatabase();
                db.Files = ViewToIEnumerable(database, "SELECT * FROM File", (file) =>
                new MsiFile
                {
                    File = file.GetString(1),
                    Component = file.GetString(2),
                    FileName = file.GetString(3),
                    FileSize = file.GetString(4),
                    Version = file.GetString(5),
                    Language = file.GetString(6),
                    Atttibutes = file.GetString(7),
                    Sequence = file.GetString(8)
                });

                db.Directories = ViewToIEnumerable(database, "SELECT * FROM Directory", (dir) =>
                new MsiDirectory
                {
                    Directory = dir.GetString(1),
                    DirectoryParent = dir.GetString(2),
                    DefaultDir = dir.GetString(3)
                });

                db.Components = ViewToIEnumerable(database, "SELECT * FROM Component", (comp) =>
                new MsiComponent
                {
                    Component = comp.GetString(1),
                    ComponentId = comp.GetString(2),
                    Directory = comp.GetString(3),
                    Condition = comp.GetString(4),
                    KeyPath = comp.GetString(5),
                    Attributes = comp.GetString(6)
                });
                return db;
            }
        }

        static IEnumerable<T> ViewToIEnumerable<T>(Database db, string sql, Func<Record, T> func)
        {
            var view = db.OpenView(sql);
            view.Execute();
            var record = view.Fetch();
            var list = new List<T>();
            while (record != null)
            {
                list.Add(func(record));
                record = view.Fetch();
            }
            return list;
        }

        static string GetFullPath(MsiDatabase db, string dirId, string pathSoFar)
        {
            var dir = db.Directories.Where(d => d.Directory == dirId).First();
            var parent = db.Directories.Where(d => d.Directory == dir.DirectoryParent).FirstOrDefault();
            var dirName = dir.DefaultDir == "." ? $"[{dir.Directory}]" : dir.DefaultDir.Split('|').Last();

            return parent != null
                ? GetFullPath(db, dir.DirectoryParent, dirName + "\\" + pathSoFar)
                : dirName + "\\" + pathSoFar;

        }
    }

    class MsiDatabase
    {
        public IEnumerable<MsiComponent> Components { get; set; }
        public IEnumerable<MsiDirectory> Directories { get; set; }
        public IEnumerable<MsiFile> Files { get; set; }
    }

    class MsiDirectory
    {
        public string Directory { get; set; }
        public string DirectoryParent { get; set; }
        public string DefaultDir { get; set; }
    }

    class MsiComponent
    {
        public string Component { get; set; }
        public string ComponentId { get; set; }
        public string Directory { get; set; }
        public string Attributes { get; set; }
        public string Condition { get; set; }
        public string KeyPath { get; set; }
    }

    class MsiFile
    {
        public string File { get; set; }
        public string Component { get; set; }
        public string FileName { get; set; }
        public string FileSize { get; set; }
        public string Version { get; set; }
        public string Language { get; set; }
        public string Atttibutes { get; set; }
        public string Sequence { get; set; }
    }
}
