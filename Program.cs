// Program.cs (contains all classes in a single file)
using Microsoft.Office.Interop.Access.Dao;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Index = Microsoft.Office.Interop.Access.Dao.Index;

namespace AccessSchemaComparer
{
    #region Models
    public class DatabaseSchema
    {
        public string? DatabasePath { get; set; }
        public List<TableSchema> Tables { get; set; } = [];
        public List<RelationSchema> Relations { get; set; } = [];
        public List<QuerySchema> Queries { get; set; } = [];
    }

    public class TableSchema
    {
        public required string Name { get; set; }
        public List<FieldSchema> Fields { get; set; } = [];
        public List<IndexSchema> Indexes { get; set; } = [];
    }

    public class FieldSchema
    {
        public required string Name { get; set; }
        public required string DataType { get; set; }
        public int Size { get; set; }
        public bool IsRequired { get; set; }
        public bool IsPrimaryKey { get; set; }
        public string? DefaultValue { get; set; }

        public override bool Equals( object? obj )
        {
            if(obj is not FieldSchema other)
                return false;

            return Name == other.Name &&
                   DataType == other.DataType &&
                   Size == other.Size &&
                   IsRequired == other.IsRequired &&
                   IsPrimaryKey == other.IsPrimaryKey &&
                   DefaultValue == other.DefaultValue;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine( Name, DataType, Size, IsRequired, IsPrimaryKey, DefaultValue );
        }
    }

    public class IndexSchema
    {
        public required string Name { get; set; }
        public bool IsPrimary { get; set; }
        public bool IsUnique { get; set; }
        public List<string> Fields { get; set; } = [];
    }

    public class RelationSchema
    {
        public required string Name { get; set; }
        public required string SourceTable { get; set; }
        public string? SourceColumn { get; set; }
        public required string TargetTable { get; set; }
        public string? TargetColumn { get; set; }

        public override bool Equals( object? obj )
        {
            if(obj is not RelationSchema other)
                return false;

            return Name == other.Name &&
                   SourceTable == other.SourceTable &&
                   SourceColumn == other.SourceColumn &&
                   TargetTable == other.TargetTable &&
                   TargetColumn == other.TargetColumn;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine( Name, SourceTable, SourceColumn, TargetTable, TargetColumn );
        }
    }

    public class QuerySchema
    {
        public required string Name { get; set; }

        public override bool Equals( object? obj )
        {
            if(obj is not QuerySchema other)
                return false;

            return Name == other.Name;
        }

        public override int GetHashCode()
        {
            return Name.GetHashCode();
        }
    }
    #endregion

    #region SchemaReader
    public static class DaoExtensions
    {
        public static string? GetDescription( this Database database )
        {
            return HasProperty( database.Properties, "Description" ) ? GetProperty( database.Properties, "Description" ) : null;
        }

        public static string? GetDescription( this Field field )
        {
            return HasProperty( field.Properties, "Description" ) ? GetProperty( field.Properties, "Description" ) : null;
        }

        public static string? GetDescription( this TableDef tableDef )
        {
            return HasProperty( tableDef.Properties, "Description" ) ? GetProperty( tableDef.Properties, "Description" ) : null;
        }

        public static IEnumerable<Field> GetFields( this Index index )
        {
            object fields = index.Fields;
            Type fieldsType = fields.GetType();

            short fieldCount = (short)(fieldsType?.InvokeMember("Count", BindingFlags.GetProperty, null, fields, null) ?? 0);

            for(short fieldIter = 0; fieldIter < fieldCount; fieldIter++)
            {
                yield return (Field)(fieldsType?.InvokeMember( "Item", BindingFlags.GetProperty, null, fields, new object[] { fieldIter } ) ?? throw new Exception());
            }
        }

        public static int GetPrecision( this Field field )
        {
            return HasProperty( field.Properties, "Precision" ) ? GetProperty<short>( field.Properties, "Precision" ) : 0;
        }

        private static T GetProperty<T>( Properties properties, string propertyName ) where T : struct
        {
            return (T)properties[propertyName].Value;
        }

        private static string GetProperty( Properties properties, string propertyName )
        {
            return (string)properties[propertyName].Value;
        }

        public static int GetScale( this Field field )
        {
            return HasProperty( field.Properties, "Scale" ) ? GetProperty<short>( field.Properties, "Scale" ) : 0;
        }

        public static bool GetUnicodeCompression( this Field field )
        {
            return HasProperty( field.Properties, "UnicodeCompression" ) ? GetProperty<bool>( field.Properties, "UnicodeCompression" ) : false;
        }

        private static bool HasProperty( Properties properties, string propertyName )
        {
            try
            {
                return properties[propertyName]?.Value != null;
            }
            catch(Exception ex)
            {
                if(ex.HResult != unchecked((int)0x800A0CC6)) throw;
                return false;
            }
        }
    }
    public class SchemaReader
    {
        // DAO Data Type Constants
        private const int dbBoolean = 1;
        private const int dbByte = 2;
        private const int dbInteger = 3;
        private const int dbLong = 4;
        private const int dbCurrency = 5;
        private const int dbSingle = 6;
        private const int dbDouble = 7;
        private const int dbDate = 8;
        private const int dbText = 10;
        private const int dbLongBinary = 11;
        private const int dbMemo = 12;
        private const int dbGUID = 15;

        // Primary Key Index Constant
        private const int dbPrimaryKey = 2;

        public DatabaseSchema ReadSchema( string databasePath )
        {
            var schema = new DatabaseSchema
            {
                DatabasePath = databasePath
            };

            // Initialize the DAO DBEngine
            DBEngine? dbEngine = null;
            Database? db = null;

            try
            {
                dbEngine = new DBEngine();
                // Open database in read-only mode (false for exclusive, true for read-only)
                db = dbEngine.OpenDatabase(databasePath, false, true);

                // Read tables, fields, and indexes
                ReadTables( db, schema );

                // Read relations
                ReadRelations( db, schema );

                // Read queries
                ReadQueries( db, schema );
            }
            finally
            {
                // Clean up COM objects to prevent memory leaks
                if(db != null)
                {
                    Marshal.ReleaseComObject( db );
                }

                if(dbEngine != null)
                {
                    Marshal.ReleaseComObject( dbEngine );
                }

                // Ensure garbage collection is triggered
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return schema;
        }

        private void ReadTables( Database db, DatabaseSchema schema )
        {
            foreach(TableDef tableDef in db.TableDefs)
            {
                // Skip system tables (those that start with "MSys")
                if(tableDef.Name.StartsWith( "MSys" ) || tableDef.Name.StartsWith( "~" ))
                {
                    continue;
                }

                var tableSchema = new TableSchema { Name = tableDef.Name };

                // Read fields
                foreach(Field field in tableDef.Fields)
                {
                    var fieldSchema = new FieldSchema
                    {
                        Name = field.Name,
                        DataType = GetDataTypeName(field.Type),
                        Size = field.Size,
                        IsRequired = field.Required,
                        DefaultValue = field.DefaultValue?.ToString()
                    };

                    tableSchema.Fields.Add( fieldSchema );
                }

                // Read indexes and determine primary key fields
                try
                {
                    for(int i = 0; i < tableDef.Indexes.Count; i++)
                    {
                        var index = tableDef.Indexes[i];

                        var indexSchema = new IndexSchema
                        {
                            Name = index.Name,
                            IsPrimary = index.Primary,
                            IsUnique = index.Unique
                        };

                        // Get fields in the index
                        var fieldsCollection = index.GetFields();
                        foreach(var indexField in fieldsCollection)
                        {
                            indexSchema.Fields.Add( indexField.Name );

                            // If this is a primary key index, mark the field as primary key
                            if(index.Primary)
                            {
                                var pkField = tableSchema.Fields.Find(f => f.Name == indexField.Name);
                                if(pkField != null)
                                {
                                    pkField.IsPrimaryKey = true;
                                }
                            }
                        }

                        tableSchema.Indexes.Add( indexSchema );
                    }
                }
                catch(Exception ex) when(0x800A0C26 == (uint)ex.HResult)
                {
                    // Handle case where a linked table doesn't support indexes
                }

                schema.Tables.Add( tableSchema );
            }
        }

        private void ReadRelations( Database db, DatabaseSchema schema )
        {
            foreach(Relation relation in db.Relations)
            {
                var relationSchema = new RelationSchema
                {
                    Name = relation.Name,
                    SourceTable = relation.ForeignTable,
                    TargetTable = relation.Table
                };

                // Get the field names from the relation
                if(relation.Fields.Count > 0)
                {
                    relationSchema.SourceColumn = relation.Fields[0].ForeignName;
                    relationSchema.TargetColumn = relation.Fields[0].Name;
                }

                schema.Relations.Add( relationSchema );
            }
        }

        private void ReadQueries( Database db, DatabaseSchema schema )
        {
            foreach(QueryDef queryDef in db.QueryDefs)
            {
                // Filter out system queries
                if(!queryDef.Name.StartsWith( "~" ))
                {
                    var querySchema = new QuerySchema { Name = queryDef.Name };
                    schema.Queries.Add( querySchema );
                }
            }
        }

        private string GetDataTypeName( int dataType )
        {
            switch(dataType)
            {
                case dbBoolean:
                    return "Boolean";
                case dbByte:
                    return "Byte";
                case dbInteger:
                    return "Integer";
                case dbLong:
                    return "Long Integer";
                case dbCurrency:
                    return "Currency";
                case dbSingle:
                    return "Single";
                case dbDouble:
                    return "Double";
                case dbDate:
                    return "Date/Time";
                case dbText:
                    return "Text";
                case dbLongBinary:
                    return "Long Binary";
                case dbMemo:
                    return "Memo";
                case dbGUID:
                    return "GUID";
                default:
                    return $"Unknown ({dataType})";
            }
        }
    }
    #endregion

    #region SchemaComparer
    public class SchemaComparer
    {
        public string Compare( DatabaseSchema schema1, DatabaseSchema schema2 )
        {
            StringBuilder output = new StringBuilder();

            output.AppendLine( "Database Schema Comparison Report" );
            output.AppendLine( "===============================" );
            output.AppendLine();
            output.AppendLine( $"First Database: {schema1.DatabasePath}" );
            output.AppendLine( $"Second Database: {schema2.DatabasePath}" );
            output.AppendLine( $"Generated: {DateTime.Now}" );
            output.AppendLine();

            // Compare tables
            CompareTables( schema1, schema2, output );

            // Compare relations
            CompareRelations( schema1, schema2, output );

            // Compare queries
            CompareQueries( schema1, schema2, output );

            return output.ToString();
        }

        private void CompareTables( DatabaseSchema schema1, DatabaseSchema schema2, StringBuilder output )
        {
            output.AppendLine( "Table Comparison" );
            output.AppendLine( "===============" );
            output.AppendLine();

            // Tables in schema1 but not in schema2
            var tablesOnlyInSchema1 = schema1.Tables
                .Where(t1 => !schema2.Tables.Any(t2 => t2.Name == t1.Name))
                .Select(t => t.Name)
                .ToList();

            if(tablesOnlyInSchema1.Any())
            {
                output.AppendLine( "Tables in first database only:" );
                foreach(var tableName in tablesOnlyInSchema1)
                {
                    output.AppendLine( $"  - {tableName}" );
                }
                output.AppendLine();
            }

            // Tables in schema2 but not in schema1
            var tablesOnlyInSchema2 = schema2.Tables
                .Where(t2 => !schema1.Tables.Any(t1 => t1.Name == t2.Name))
                .Select(t => t.Name)
                .ToList();

            if(tablesOnlyInSchema2.Any())
            {
                output.AppendLine( "Tables in second database only:" );
                foreach(var tableName in tablesOnlyInSchema2)
                {
                    output.AppendLine( $"  - {tableName}" );
                }
                output.AppendLine();
            }

            // Compare fields in tables that exist in both schemas
            foreach(var table1 in schema1.Tables)
            {
                var table2 = schema2.Tables.FirstOrDefault(t => t.Name == table1.Name);
                if(table2 != null)
                {
                    CompareTableFields( table1, table2, output );
                    CompareTableIndexes( table1, table2, output );
                }
            }
        }

        private void CompareTableFields( TableSchema table1, TableSchema table2, StringBuilder output )
        {
            bool hasDifferences = false;

            // Fields in table1 but not in table2
            var fieldsOnlyInTable1 = table1.Fields
                .Where(f1 => !table2.Fields.Any(f2 => f2.Name == f1.Name))
                .ToList();

            // Fields in table2 but not in table1
            var fieldsOnlyInTable2 = table2.Fields
                .Where(f2 => !table1.Fields.Any(f1 => f1.Name == f2.Name))
                .ToList();

            // Fields with different properties
            var commonFields = table1.Fields
                .Where(f1 => table2.Fields.Any(f2 => f2.Name == f1.Name))
                .ToList();

            var fieldDifferences = new List<(FieldSchema Field1, FieldSchema Field2)>();
            foreach(var field1 in commonFields)
            {
                var field2 = table2.Fields.First(f => f.Name == field1.Name);
                if(!field1.Equals( field2 ))
                {
                    fieldDifferences.Add( (field1, field2) );
                }
            }

            // Output differences if any
            if(fieldsOnlyInTable1.Any() || fieldsOnlyInTable2.Any() || fieldDifferences.Any())
            {
                output.AppendLine( $"Differences in table '{table1.Name}':" );
                hasDifferences = true;

                if(fieldsOnlyInTable1.Any())
                {
                    output.AppendLine( "  Fields in first database only:" );
                    foreach(var field in fieldsOnlyInTable1)
                    {
                        output.AppendLine( $"    - {field.Name} ({field.DataType}, Size: {field.Size})" );
                    }
                }

                if(fieldsOnlyInTable2.Any())
                {
                    output.AppendLine( "  Fields in second database only:" );
                    foreach(var field in fieldsOnlyInTable2)
                    {
                        output.AppendLine( $"    - {field.Name} ({field.DataType}, Size: {field.Size})" );
                    }
                }

                if(fieldDifferences.Any())
                {
                    output.AppendLine( "  Fields with different properties:" );
                    foreach(var (field1, field2) in fieldDifferences)
                    {
                        output.AppendLine( $"    - {field1.Name}:" );

                        if(field1.DataType != field2.DataType)
                            output.AppendLine( $"      DataType: {field1.DataType} => {field2.DataType}" );

                        if(field1.Size != field2.Size)
                            output.AppendLine( $"      Size: {field1.Size} => {field2.Size}" );

                        if(field1.IsRequired != field2.IsRequired)
                            output.AppendLine( $"      Required: {field1.IsRequired} => {field2.IsRequired}" );

                        if(field1.IsPrimaryKey != field2.IsPrimaryKey)
                            output.AppendLine( $"      PrimaryKey: {field1.IsPrimaryKey} => {field2.IsPrimaryKey}" );

                        if(field1.DefaultValue != field2.DefaultValue)
                            output.AppendLine( $"      DefaultValue: {field1.DefaultValue} => {field2.DefaultValue}" );
                    }
                }
            }

            if(hasDifferences)
            {
                output.AppendLine();
            }
        }

        private void CompareTableIndexes( TableSchema table1, TableSchema table2, StringBuilder output )
        {
            // Indexes in table1 but not in table2
            var indexesOnlyInTable1 = table1.Indexes
                .Where(i1 => !table2.Indexes.Any(i2 => i2.Name == i1.Name))
                .ToList();

            // Indexes in table2 but not in table1
            var indexesOnlyInTable2 = table2.Indexes
                .Where(i2 => !table1.Indexes.Any(i1 => i1.Name == i2.Name))
                .ToList();

            if(indexesOnlyInTable1.Any() || indexesOnlyInTable2.Any())
            {
                output.AppendLine( $"Index differences in table '{table1.Name}':" );

                if(indexesOnlyInTable1.Any())
                {
                    output.AppendLine( "  Indexes in first database only:" );
                    foreach(var index in indexesOnlyInTable1)
                    {
                        output.AppendLine( $"    - {index.Name} (Fields: {string.Join( ", ", index.Fields )})" );
                    }
                }

                if(indexesOnlyInTable2.Any())
                {
                    output.AppendLine( "  Indexes in second database only:" );
                    foreach(var index in indexesOnlyInTable2)
                    {
                        output.AppendLine( $"    - {index.Name} (Fields: {string.Join( ", ", index.Fields )})" );
                    }
                }

                output.AppendLine();
            }
        }

        private void CompareRelations( DatabaseSchema schema1, DatabaseSchema schema2, StringBuilder output )
        {
            output.AppendLine( "Relation Comparison" );
            output.AppendLine( "==================" );
            output.AppendLine();

            // Relations in schema1 but not in schema2
            var relationsOnlyInSchema1 = schema1.Relations
                .Where(r1 => !schema2.Relations.Any(r2 =>
                    r2.SourceTable == r1.SourceTable &&
                    r2.SourceColumn == r1.SourceColumn &&
                    r2.TargetTable == r1.TargetTable &&
                    r2.TargetColumn == r1.TargetColumn))
                .ToList();

            if(relationsOnlyInSchema1.Any())
            {
                output.AppendLine( "Relations in first database only:" );
                foreach(var relation in relationsOnlyInSchema1)
                {
                    output.AppendLine( $"  - {relation.Name}: {relation.SourceTable}.{relation.SourceColumn} -> {relation.TargetTable}.{relation.TargetColumn}" );
                }
                output.AppendLine();
            }

            // Relations in schema2 but not in schema1
            var relationsOnlyInSchema2 = schema2.Relations
                .Where(r2 => !schema1.Relations.Any(r1 =>
                    r1.SourceTable == r2.SourceTable &&
                    r1.SourceColumn == r2.SourceColumn &&
                    r1.TargetTable == r2.TargetTable &&
                    r1.TargetColumn == r2.TargetColumn))
                .ToList();

            if(relationsOnlyInSchema2.Any())
            {
                output.AppendLine( "Relations in second database only:" );
                foreach(var relation in relationsOnlyInSchema2)
                {
                    output.AppendLine( $"  - {relation.Name}: {relation.SourceTable}.{relation.SourceColumn} -> {relation.TargetTable}.{relation.TargetColumn}" );
                }
                output.AppendLine();
            }

            if(!relationsOnlyInSchema1.Any() && !relationsOnlyInSchema2.Any())
            {
                output.AppendLine( "No differences found in relations" );
                output.AppendLine();
            }
        }

        private void CompareQueries( DatabaseSchema schema1, DatabaseSchema schema2, StringBuilder output )
        {
            output.AppendLine( "Query Comparison" );
            output.AppendLine( "===============" );
            output.AppendLine();

            // Queries in schema1 but not in schema2
            var queriesOnlyInSchema1 = schema1.Queries
                .Where(q1 => !schema2.Queries.Any(q2 => q2.Name == q1.Name))
                .Select(q => q.Name)
                .ToList();

            if(queriesOnlyInSchema1.Any())
            {
                output.AppendLine( "Queries in first database only:" );
                foreach(var queryName in queriesOnlyInSchema1)
                {
                    output.AppendLine( $"  - {queryName}" );
                }
                output.AppendLine();
            }

            // Queries in schema2 but not in schema1
            var queriesOnlyInSchema2 = schema2.Queries
                .Where(q2 => !schema1.Queries.Any(q1 => q1.Name == q2.Name))
                .Select(q => q.Name)
                .ToList();

            if(queriesOnlyInSchema2.Any())
            {
                output.AppendLine( "Queries in second database only:" );
                foreach(var queryName in queriesOnlyInSchema2)
                {
                    output.AppendLine( $"  - {queryName}" );
                }
                output.AppendLine();
            }

            if(!queriesOnlyInSchema1.Any() && !queriesOnlyInSchema2.Any())
            {
                output.AppendLine( "No differences found in queries" );
                output.AppendLine();
            }
        }
    }
    #endregion

    #region Program
    public class Program
    {
        static async Task Main( string[] args )
        {
            Console.WriteLine( "Access Database Schema Comparison Tool" );
            Console.WriteLine( "--------------------------------------" );

            string? sourcePath1, sourcePath2, outputPath;

            // Handle command line arguments or prompt for inputs
            if(args.Length >= 3)
            {
                sourcePath1 = args[0];
                sourcePath2 = args[1];
                outputPath = args[2];
            }
            else
            {
                Console.Write( "Enter path to first Access database: " );
                sourcePath1 = Console.ReadLine();

                Console.Write( "Enter path to second Access database: " );
                sourcePath2 = Console.ReadLine();

                Console.Write( "Enter output text file path: " );
                outputPath = Console.ReadLine();
            }

            // Validate input files exist
            if(!File.Exists( sourcePath1 ))
            {
                Console.WriteLine( $"Error: First database file not found: {sourcePath1}" );
                return;
            }

            if(!File.Exists( sourcePath2 ))
            {
                Console.WriteLine( $"Error: Second database file not found: {sourcePath2}" );
                return;
            }

            if(outputPath is null)
            {
                Console.WriteLine( $"Error: Output path not specified" );
                return;
            }

            try
            {
                // Create schema reader
                var reader = new SchemaReader();

                Console.WriteLine( "Reading schema from first database..." );
                var schema1 = await Task.Run(() => reader.ReadSchema(sourcePath1));

                Console.WriteLine( "Reading schema from second database..." );
                var schema2 = await Task.Run(() => reader.ReadSchema(sourcePath2));

                Console.WriteLine( "Comparing schemas..." );
                var comparer = new SchemaComparer();
                var differences = comparer.Compare(schema1, schema2);

                Console.WriteLine( "Writing differences to output file..." );
                await File.WriteAllTextAsync( outputPath, differences );

                Console.WriteLine( $"Comparison complete. Results written to {outputPath}" );
            }
            catch(Exception ex)
            {
                Console.WriteLine( $"Error: {ex.Message}" );
                if(ex.InnerException != null)
                {
                    Console.WriteLine( $"Inner Error: {ex.InnerException.Message}" );
                }
            }

            Console.WriteLine( "Press any key to exit..." );
            Console.ReadKey();
        }
    }
    #endregion
}
