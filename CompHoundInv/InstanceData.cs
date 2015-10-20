#region Namespaces
using System;
using Inventor;
using System.Diagnostics;
#endregion

namespace CompHoundInv
{
  /// <summary>
  /// Container for the family instance data to store 
  /// in the external database for the given component.
  /// </summary>
  class InstanceData
  {
    public string _id { get; set; } // : UniqueId // suppress automatic generation
    public string project { get; set; }
    public string path { get; set; }
    public string urn { get; set; } // populated later
    public string family { get; set; }
    public string symbol { get; set; }
    public string category { get; set; }
    public string level { get; set; }
    public int x { get; set; }
    public int y { get; set; }
    public int z { get; set; }
    public double easting { get; set; }
    public double northing { get; set; }
    public string properties { get; set; } // : String // json dictionary of instance properties and values

    public InstanceData()
    {
    }

    public InstanceData(
      ComponentOccurrence occ, int index )
    {
      // Could be assembly or part, but their relevant 
      // properties are the same, so can use "dynamic"
      dynamic doc = occ.Definition.Document;

      AssemblyDocument asm = occ.ContextDefinition.Document as AssemblyDocument;

      // occurrences have no unique ID, only documents have
      // let's create it based on doc ID + component index
      string internalName = doc.InternalName + "-" + index.ToString(); // Returns GUID
      Debug.Print(internalName);

      Vector pos = occ.Transformation.Translation;  
      string jsonProps = Util.GetPropertiesJson( );

      // /a/src/web/CompHoundWeb/model/instance.js
      // _id        : UniqueId // suppress automatic generation
      // project    : String
      // path       : String
      // urn        : String
      // family     : String
      // symbol     : String
      // category   : String
      // level      : String
      // x          : Number
      // y          : Number
      // z          : Number
      // easting    : Number // Geo2d?
      // northing   : Number
      // properties : String // json dictionary of instance properties and values

      _id = internalName; 
      project = asm.DisplayName; // Assembly doc name
      path = asm.File.FullFileName; // Assembly doc path
      urn = string.Empty;
      family = doc.DisplayName; 
      symbol = occ.Name; // Component occurrence name?
      category = "";
      level = "-1";
      x = Util.ConvertCmToMillimetres( pos.X );
      y = Util.ConvertCmToMillimetres( pos.Y );
      z = Util.ConvertCmToMillimetres( pos.Z );
      easting = 0;
      northing = 0; 
      properties = jsonProps;
    }
  }
}
