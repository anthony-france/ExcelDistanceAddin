using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Collections;

namespace ACF_DistanceFinder2012
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnCalculateDistances_Click(object sender, RibbonControlEventArgs e)
        {

            // Setup ranges.  
            Range rngSource;
            Range rngDest;
            Range rngOutput;
            // Probably a better idea to instead have the user setup named ranges.

            // Prompt the user to select each of our required points.
            // If it fails tell the user and return method.
            try
            {
                rngSource = Globals.ThisAddIn.Application.InputBox("Source Range", "Source", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                rngDest = Globals.ThisAddIn.Application.InputBox("Destination Columns", "Destinations", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                rngOutput = Globals.ThisAddIn.Application.InputBox("Output Starting Cell", "Output", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            }
            catch
            {
                MessageBox.Show("Error getting input.");
                return;
            }

            // counters for filling in the output cells.
            int rowOffset = 1;
            int colOffset = 0;

            // List of the destination headers.
           List<string> Destinations = new List<string>();

            // Add the colums to the Destination list to use later.
            // If it fails we need to let the user know and stop.
           try
           {
               foreach (Range rowDest in rngDest.Columns)
               {
                   Destinations.Add(System.Convert.ToString(rowDest.Value2));
               }
           }
           catch
           {
              MessageBox.Show("Unable to get destinations.");
              return;
           }

            // Loop through the list of source locations.
            foreach (Range sourceCell in rngSource.Rows)
            {
                // declare in the scope for using later.
                List<double> distance = new List<double>();
                string json;

                // Setup the request url.  
                //Better to use the uribuilder but speed isn't much of a concern.
                string url = "http://maps.googleapis.com/maps/api/distancematrix/json?sensor=false";
                url += "&origins=" + System.Convert.ToString(sourceCell.Value2);
                url += "&destinations=";

                foreach (string Destination in Destinations)
                {
                    url += Destination.ToString() + "|";
                }

                // Create the webrequest.
                // Try the next row if it fails; continue on the loop.
                try
                {
                    WebRequest request = WebRequest.Create(url);
                    WebResponse response = request.GetResponse();

                    Stream dataStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(dataStream);
                    json = reader.ReadToEnd();
                }
                catch {
                    //MessageBox.Show("Error fetching data.");
                    continue;
                }

                // Create list of returned values.
                // Continue on if fails.
                try{
                    JObject jObject = JObject.Parse(json);
                    
                    foreach (JObject rows in jObject["rows"].Children())
                    {
                        foreach (JObject elements in rows["elements"].Children())
                        {
                            distance.Add((double)elements["distance"]["value"] * 0.00062137119);
                         }
                    }
                }
                catch
                {
                    //MessageBox.Show("Error parsing data.");
                    continue;
                }

                // Place the values.
                // If failed continue on.
                try
                {
                    colOffset = 0;
                    foreach (double d in distance)
                    {
                        rngOutput[rowOffset, colOffset + 1].Value = distance[colOffset];
                        rngOutput[rowOffset, colOffset + 1].Show();
                        
                        colOffset++;
                    }
                }
                catch
                {
                    //MessageBox.Show("Problem setting data in output.");
                    continue;
                }
                rowOffset++;
                System.Threading.Thread.Sleep(500);
            }
        }

    }
}
