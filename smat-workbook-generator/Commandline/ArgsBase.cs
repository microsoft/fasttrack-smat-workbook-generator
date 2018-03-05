using Microsoft.FastTrack.SMATWorkbookGenerator.Localization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Microsoft.FastTrack.SMATWorkbookGenerator.Commandline
{
    /// <summary>
    /// Base class for argument parsing
    /// </summary>
    class ArgsBase
    {
        /// <summary>
        /// The set of property/attribute pairs used to populate this instance
        /// </summary>
        private KeyValuePair<PropertyInfo, ArgPropAttribute>[] _pairs;

        /// <summary>
        /// Creates a new instance of the Arguments, loading the properties from the supplied args
        /// </summary>
        /// <param name="args">User supplied args</param>
        public ArgsBase(string[] args)
        {
            this._pairs = null;
            this.HelpFlag = false;
            this.LoadProps(args);
        }

        /// <summary>
        /// Indicates if the user has requested help
        /// </summary>
        public bool HelpFlag { get; protected set; }

        /// <summary>
        /// Loads up this instances properties using reflection
        /// </summary>
        /// <param name="args">User supplied args</param>
        protected virtual void LoadProps(string[] args)
        {
            // if they have asked for help, flag that and the calling program is responsible for calling the WriteHelp method.
            if (args.Any((s) => Regex.IsMatch(s, Strings.HelpFlagRegEx, RegexOptions.IgnoreCase | RegexOptions.Compiled)))
            {
                this.HelpFlag = true;
                return;
            }

            // create a map to help lookup the props
            var map = new Dictionary<string, string>();

            for (var i = 0; i < args.Length; i++)
            {
                map.Add(args[i].ToLower(), args[i + 1]);
                i++;
            }

            foreach (var pair in this.PropAttrPairs)
            {
                var candidateValue = map.Keys.Any(s => s.Equals(pair.Value.Name)) ? map[pair.Value.Name] : null;
                if (candidateValue == null)
                {
                    if (pair.Value.Required)
                    {
                        throw new ArgumentNullException(string.Format(Strings.RequiredPropNotFoundExceptionMsg, pair.Value.Name));
                    }

                    candidateValue = pair.Value.Default;
                }

                pair.Key.SetValue(this, candidateValue, null);
            }
        }

        /// <summary>
        /// Array of kvp's with a PropertyInfo key and ArgProp value
        /// </summary>
        protected virtual KeyValuePair<PropertyInfo, ArgPropAttribute>[] PropAttrPairs
        {
            get
            {

                if (this._pairs == null)
                {
                    var pairs = new List<KeyValuePair<PropertyInfo, ArgPropAttribute>>();
                    foreach (var prop in this.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
                    {
                        var attr = prop.GetCustomAttributes(typeof(ArgPropAttribute), true).Cast<ArgPropAttribute>().FirstOrDefault();

                        if (attr == null)
                        {
                            continue;
                        }

                        pairs.Add(new KeyValuePair<PropertyInfo, ArgPropAttribute>(prop, attr));
                    };

                    this._pairs = pairs.ToArray();
                }

                return this._pairs;
            }
        }

        /// <summary>
        /// Writes the help for this argument parser
        /// </summary>
        /// <param name="writer">The writer to which we will output the help messages</param>
        public virtual void WriteHelp(string appTitle, string appDescription, TextWriter writer)
        {
            writer.WriteLine(appTitle);
            writer.WriteLine(appDescription);
            writer.WriteLine();

            foreach (var pair in this.PropAttrPairs)
            {
                if (!pair.Value.Required)
                {
                    writer.Write("{0} ", Strings.HelpOptionalText);
                }

                writer.WriteLine("{0} : {1}", pair.Value.Name, pair.Value.Description);

                if (!pair.Value.Required)
                {
                    writer.WriteLine("{0} {1}", Strings.HelpDefaultValueText, pair.Value.Default);
                }

                writer.WriteLine();
            }
        }
    }

    /// <summary>
    /// Defines the property used to tag argument parser class's properties
    /// </summary>
    class ArgPropAttribute : Attribute
    {
        /// <summary>
        /// Creates a new instance of the ArgPropAttribute class for a required argument
        /// </summary>
        /// <param name="name">Name of the argument on the command line (what someone would type)</param>
        /// <param name="description">A description output for help</param>
        public ArgPropAttribute(string name, string description)
        {
            this.Name = name.ToLower();
            this.Description = description;
            this.Required = true;
            this.Default = null;
        }

        /// <summary>
        /// Creates a new instance of the ArgPropAttribute class for a non-required argument
        /// </summary>
        /// <param name="name">Name of the argument on the command line (what someone would type)</param>
        /// <param name="description">A description output for help</param>
        /// <param name="default">Default value for the argument if it is not supplied by the user</param>
        public ArgPropAttribute(string name, string description, string @default)
        {
            this.Name = name.ToLower();
            this.Description = description;
            this.Required = false;
            this.Default = @default;
        }

        /// <summary>
        /// Name of the argument on the command line (what someone would type)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// A description output for help
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Indicates if this argument is required
        /// </summary>
        public bool Required { get; set; }

        /// <summary>
        /// Default argument value, if available
        /// </summary>
        public string Default { get; set; }
    }
}
