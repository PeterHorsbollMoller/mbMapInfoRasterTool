using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using MapInfo.RasterEngine.Common;
using MapInfo.RasterEngine.Operations;
using MapInfo.RasterEngine.IO;
using MapInfo.Types;

namespace MapInfoProAdvancedTool
{
    public interface IMapInfoProAddInInterface
    {
        void Initialize(IMapInfoPro mapInfoApplication, string mbxname);
        void Unload();
        IMapBasicApplication ThisApplication { get; set; }

    }

    public class MapInfoProAddInBase : IMapInfoProAddInInterface
    {
        protected IMapInfoPro MapInfoApplication;

        public virtual void Initialize(IMapInfoPro mapInfoApplication, string mbxname)
        {
            //UriParser.Register(new GenericUriParser(GenericUriParserOptions.GenericAuthority), "pack", -1);
            MapInfoApplication = mapInfoApplication;
        }

        public virtual void Unload()
        {
        }

        public IMapBasicApplication ThisApplication { get; set; }

        public static Uri PathToUri(string uri)
        {
            try
            {
                return new Uri(uri, UriKind.RelativeOrAbsolute);
            }
            catch (Exception)
            {

                return null;
            }
        }
    }

    public class ProAddIn : MapInfoProAddInBase
    {
        #region private members
        private const string AddinName = "Pro Raster AddIn";

        private string _uniqueIdentifier; //Unique name to identify running addin 
        private string _applicationDirectory = string.Empty;

        #endregion

        /// <summary>
        /// Initalize the addin
        /// </summary>
        /// <param name="mapInfoApplication">Application instance</param>
        /// <param name="mbxname">Addin name</param>
        public override void Initialize(IMapInfoPro mapInfoApplication, string mbxname)
        {
            mapInfoApplication.LogEvent(AddinName, "StartUp", null);
            // sets MapInfoApplication and ThisApplication
            base.Initialize(mapInfoApplication, mbxname);

            //_applicationDirectory = _application.EvalMapBasicCommand("ApplicationDirectory$()");
            _uniqueIdentifier = mbxname.Split('.')[0];
        }

        /// <summary>
        /// Reset the customization
        /// </summary>
        public override void Unload()
        {

        }
    }

        public class RasterTools
    {
        //// <code>
        public static string MergeAsMRR(string[] strInFilePaths, string strOutFilePath)
        {

            RasterFinalizationOptions finalizationOptions = new RasterFinalizationOptions
            {
                BuildOverviews = BuildOverviewOption.Always,
                ComputeStatistics = ComputeStatisticsOption.Always,
                StatisticsMode = RasterStatisticsMode.Advanced
            };            
            var apiOptions = new RasterApiOptions(null, finalizationOptions, null, null);

            //Merge input grids with operator
            RasterProcessing.Merge(strInFilePaths, 0, strOutFilePath, "MI_MRR", MergeOperator.Average
                                , MergeType.Union, RasterResampleMethod.Bilinear, false
                                , MergeMultiResolutionMode.OptimumMaximum, null, apiOptions, null);

            return strOutFilePath;
        }

        //// <code>
        public static string MergeAs(string[] strInFilePaths, string strOutFilePath, string rasterDriverID)
        {
            RasterFinalizationOptions finalizationOptions = new RasterFinalizationOptions
            {
                BuildOverviews = BuildOverviewOption.Always,
                ComputeStatistics = ComputeStatisticsOption.Always,
                StatisticsMode = RasterStatisticsMode.Advanced
            };
            var apiOptions = new RasterApiOptions(null, finalizationOptions, null, null);
            
            //Merge input grids
            RasterProcessing.Merge(strInFilePaths, 0, strOutFilePath, rasterDriverID, MergeOperator.Average
                                , MergeType.Union, RasterResampleMethod.Bilinear, false
                                , MergeMultiResolutionMode.OptimumMaximum, null, apiOptions, null);

            return strOutFilePath;
        }

        public static void CreateTabFileForRasterFile(string strRasterFile, string sCoordsys, bool bOverwriteTabFile)
        {
            RasterDatasetFactory.CreateTABFile(strRasterFile, sCoordsys, bOverwriteTabFile);
        }

        public static void CreateTabFilesForRasterFiles(string[] strRasterFiles, string sCoordsys, bool bOverwriteTabFile)
        {
            if (strRasterFiles.Length > 0)
            {
                foreach (string sRasterFile in strRasterFiles)
                {
                    RasterDatasetFactory.CreateTABFile(sRasterFile, sCoordsys, bOverwriteTabFile);
                }
            }
        }

        public static double GetCellAreaAboveThreshold(string inputFile, int fieldIndex, int bandIndex, double threshold)
        {
            // number of cells over threshold
            double cellCount = 0;
            // open the file and get the dataset
            var dataset = RasterDatasetFactory.Open(inputFile);
            // get the raster info
            var info = dataset.Info;
            // get sequential tile iterator for read for current field
            var iterator = dataset.GetSequentialTileIterator(fieldIndex);
            // begin the iterator
            iterator.Begin();
            // iterate over each tile in raster
            foreach (var tile in iterator)
            {
                // get band to read
                uint ubandIndex = (uint)bandIndex;
                var band = tile.GetBandData<double>(ubandIndex);
                // iterate over each cell in tile and read
                for (uint y = 0; y < tile.YSize; y++)
                {
                    for (uint x = 0; x < tile.XSize; x++)
                    {
                        double value;
                        bool valid;
                        // get the cell value
                        band.GetCellValue(x + y * tile.XSize, out value, out valid);
                        // check its valid and above threshold
                        if (valid && (value > threshold))
                            cellCount++;
                    }
                }
            }
            // end the iterator
            iterator.End();
            // close the dataset
            dataset.Close();
            // calculate area
            return cellCount * (info.FieldInfo[fieldIndex].CellSizeX.Decimal * info.FieldInfo[fieldIndex].CellSizeY.Decimal);
        }

        public static string ContinuousToClassifiedMRR(string strInFilePath, string strOutFilePath)
        {
            const ClassificationType outClassType = ClassificationType.Classified;

            UndefinedClassInfo undefinedClassInfo = new UndefinedClassInfo();

            undefinedClassInfo.SetNullForUndefinedRange = false;
            undefinedClassInfo.UndefinedClassColor = Color.FromArgb(0, 128, 64, 0);
            undefinedClassInfo.UndefinedClassLabel = "Undefined";
            undefinedClassInfo.UndefinedValue = -9999;

            ClassInfo[] newClassInfo = new ClassInfo[5];
            //Initalize class items
            for (int i = 0; i < newClassInfo.Length; i++)
                newClassInfo[i] = new ClassInfo();

            //Set classify ranges
            newClassInfo[0].LowerBound = 0.0f;
            newClassInfo[0].UpperBound = 4.0f;
            newClassInfo[1].LowerBound = 4.0f;
            newClassInfo[1].UpperBound = 8.0f;
            newClassInfo[2].LowerBound = 8.0f;
            newClassInfo[2].UpperBound = 12.0f;
            newClassInfo[3].LowerBound = 12.0f;
            newClassInfo[3].UpperBound = 16.0f;
            newClassInfo[4].LowerBound = 16.0f;
            newClassInfo[4].UpperBound = 20.0f;

            //Set new class values
            newClassInfo[0].NewValue = 4.0f;
            newClassInfo[1].NewValue = 8.0f;
            newClassInfo[2].NewValue = 12.0f;
            newClassInfo[3].NewValue = 16.0f;
            newClassInfo[4].NewValue = 20.0f;

            //Set new class names
            newClassInfo[0].NewLabel = "0 - 4";
            newClassInfo[1].NewLabel = "4 - 8";
            newClassInfo[2].NewLabel = "8 - 12";
            newClassInfo[3].NewLabel = "12 -16";
            newClassInfo[4].NewLabel = "16 - 20";

            //Set new class color
            newClassInfo[0].NewColor = Color.FromArgb(0, 0, 255, 255);//cyan
            newClassInfo[1].NewColor = Color.FromArgb(0, 255, 128, 192);//Pink Color
            newClassInfo[2].NewColor = Color.FromArgb(0, 255, 128, 64);//Orange
            newClassInfo[3].NewColor = Color.FromArgb(0, 128, 0, 255);
            newClassInfo[4].NewColor = Color.FromArgb(0, 255, 0, 0);

            RasterAnalysis.ClassifyRaster(strInFilePath, strOutFilePath, "MI_MRR", outClassType, newClassInfo, undefinedClassInfo);

            return strOutFilePath;
        }

        public static string RegionInspection(    string strInTABFilePath
                                                , string[] strInRasterFilesPath
                                                , string strOutTABFilePath
                                                , bool bMinimumValue
                                                , bool bMaximumValue
                                                , bool bAverageValue
                                                , bool bMedianValue
                                                , bool bNumCells
                                                , bool bNumNullCells
                                                , bool bCoefficientOfVariance
                                                , bool bRange
                                                , bool bStandardDeviation
                                                , bool bSumOfCells
                                                , bool bLowerQuartile
                                                , bool bUpperQuartile
                                                , bool bInterQuartileRange
                                                )
        {

            PolygonStatisticFlags inStatsFlag = new PolygonStatisticFlags
            {
                Min = bMinimumValue,
                Max = bMaximumValue,
                Mean = bAverageValue,
                Median = bMedianValue,
                TotalCells = bNumCells,
                TotalNullCells = bNumNullCells,
                CoefficientOfVariance = bCoefficientOfVariance,
                Range = bRange,
                StandardDeviation = bStandardDeviation,
                SumOfCells = bSumOfCells,
                LowerQuartile = bLowerQuartile,
                UpperQuartile = bUpperQuartile,
                InterQuartileRange = bInterQuartileRange
            };

            RasterAnalysis.GetPolygonStatistics(strInTABFilePath, strInRasterFilesPath, strOutTABFilePath, inStatsFlag);

            return strOutTABFilePath;

        }

        public static string RegionInspectionBands(string strInTABFilePath
                                                , string[] strInRasterFilesPath
                                                , int fieldIndex
                                                , int bandIndex
                                                , string strOutTABFilePath
                                                , bool bMinimumValue
                                                , bool bMaximumValue
                                                , bool bAverageValue
                                                , bool bMedianValue
                                                , bool bNumCells
                                                , bool bNumNullCells
                                                , bool bCoefficientOfVariance
                                                , bool bRange
                                                , bool bStandardDeviation
                                                , bool bSumOfCells
                                                , bool bLowerQuartile
                                                , bool bUpperQuartile
                                                , bool bInterQuartileRange
                                                )
        {

            PolygonStatisticFlags inStatsFlag = new PolygonStatisticFlags
            {
                Min = bMinimumValue,
                Max = bMaximumValue,
                Mean = bAverageValue,
                Median = bMedianValue,
                TotalCells = bNumCells,
                TotalNullCells = bNumNullCells,
                CoefficientOfVariance = bCoefficientOfVariance,
                Range = bRange,
                StandardDeviation = bStandardDeviation,
                SumOfCells = bSumOfCells,
                LowerQuartile = bLowerQuartile,
                UpperQuartile = bUpperQuartile,
                InterQuartileRange = bInterQuartileRange
            };

            RasterAnalysis.GetPolygonStatistics(strInTABFilePath, strInRasterFilesPath, strOutTABFilePath, inStatsFlag, (uint)fieldIndex, (uint)bandIndex);

            return strOutTABFilePath;

        }

        public static string RasterExportBand(    string strInRasterFilePath
                                                 , int fieldIndex
                                                 , int bandIndex
                                                 , string strOutRasterFilePath
                                                 , string strOutputRasterDriver
                                                 )
        {

            // Specify FieldBandFilter so that selected Field's AllBands will be converted to form single field output raster
            // depending upon Driver capabilities
            var fieldBandFilter = new FieldBandFilter((uint)fieldIndex, (uint)bandIndex);
            //Setting RasterCreationOptions and RasterFinalizationOptions as null. This will result in default settings read from user's preference file.
            var apiOptions = new RasterApiOptions(null, null, fieldBandFilter, null);
            RasterProcessing.Convert(strInRasterFilePath, strOutRasterFilePath, strOutputRasterDriver, apiOptions);
            
             return strOutRasterFilePath;
        }

        public static void RasterDelete(string strInRasterFilePath
                                                 )
        {
            RasterProcessing.Delete(strInRasterFilePath);
        }

        public static void StampFixedGridSize( string strInputFile
                                            , string strCoordSystem
                                            , string strOutputFile
                                            , string strOutputRasterDriver
                                            , double OriginX
                                            , double OriginY
                                            , double GridSizeX
                                            , double GridSizeY
                                            , int numColumnsX
                                            , int numRowsY
                                            , int dataFieldIndex
                                            )
        {
            RasterInterpolationOptions interpolationOptions = new RasterInterpolationOptions
            {
                Preferences =
                {
                    AutoCacheSize = true,
                    CacheSize = 0,
                    TemporaryDirectory = ""
                }
            };

            // setup source data file parameters, can add more than one file
            RasterInterpolationInputFileDetails file = new RasterInterpolationInputFileDetails();
            file.DataFieldIndexes.Add((uint)dataFieldIndex);
            file.XFieldIndex = -1;
            file.YFieldIndex = -1;
            file.FileType = RasterInterpolationInputFileType.MAPINFO_TAB;
            file.FilePath = strInputFile;
            file.HeaderRows = 0;
            file.CoordSysString = strCoordSystem;
            file.SubFileName = "*.*";
            interpolationOptions.InputDetails.Add(file);

            // setup output raster parameters
            interpolationOptions.OutputDetails.FilePath = strOutputFile;
            interpolationOptions.OutputDetails.DriverId = strOutputRasterDriver;
            interpolationOptions.OutputDetails.AutoDataType = true;
            /* not used in this example as we have set AutoDataType to true
            interpolationOptions.OutputFile.DataTypes.Add(RasterBandDataType.RealSingle);		
            */
            interpolationOptions.OutputDetails.CoordSysString = strCoordSystem;

            // setup raster geometry paramters
            interpolationOptions.Geometry.Extent.Auto = false;
            // used in this example as we have set AutoGridExtents to false
            interpolationOptions.Geometry.Extent.OriginX = OriginX;
            interpolationOptions.Geometry.Extent.OriginY = OriginY;
            //interpolationOptions.Geometry.Extent.Type = RasterInterpolationExtentType.RowsColumns;
            interpolationOptions.Geometry.Extent.Type = RasterInterpolationExtentType.Cells;
            interpolationOptions.Geometry.Extent.ExtentX = numColumnsX;
            interpolationOptions.Geometry.Extent.ExtentY = numRowsY;
            
            interpolationOptions.Geometry.CellSize.Auto = false;
            // used in this example as we have set AutoGridCellSize to false
            interpolationOptions.Geometry.CellSize.X = GridSizeX;
            interpolationOptions.Geometry.CellSize.Y = GridSizeY;

            // setup stamp method parameters
            RasterInterpolationStamp stampOptions = new RasterInterpolationStamp
            {
                StampMethod = RasterInterpolationStampStampingMethod.AverageLastInWeighted,
                Clip = new RasterInterpolationClip
                {
                    Method = RasterInterpolationClipMethod.None,
                    Near = 0,
                    Far = 0,
                    KeepWithinPolygon = true,
                    PolygonFile = ""
                },
                Smoothing = new RasterInterpolationSmoothing
                {
                    Type = RasterInterpolationSmoothingType.None,
                    Level = 1
                }
            };

            RasterApiOptions apiOptions = new RasterApiOptions();

            using (IRasterProgressTracker tracker = RasterProgressFactory.Create(ProgressTrackerCallBack))
            {
                RasterInterpolation.InterpolateStamp(interpolationOptions, stampOptions, apiOptions, tracker);
            }
        }

        public static void ExportToASCII(string strInFilePath, string strOutFilePath)
        {
            // This example Export a continuous raster in to text file with all default parameters
            // FieldBandFilter is set to default. i.e. process all fields and bands
            // RasterExportExOptions is set to default
            // This sets Null Value to -9999.0
            // Writes cells from bottom left of the raster.
            // considers cell value at the cell center
            // does write X Y values in file
            // Precision is set to 6 digits after decimal
            // Delimiter to sepparate values is set to  comma i.e. ','
            RasterExportExOptions ExportOptions = new RasterExportExOptions(-9999, 6, true, false, false, ',');
            RasterProcessing.ExportRasterToASCII(strInFilePath, strOutFilePath, ExportOptions);

            //Sets FieldBandFilter to 0,0 which means it works only with field 0, Band 0
            // Sets Clip extents so that Raster is read only within specified extents
            //RasterApiOptions ApiOptions = new RasterApiOptions(null, null, new FieldBandFilter(0, 0), new RasterClipExtent(455.0, 200.0, 1000.0, 800.0));
            //RasterProcessing.ExportRasterToASCII(strInFilePath, strOutFilePath, ExportOptions, ApiOptions);
        }

        public static void RunRasterProcessorUI()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = false;
            dlg.Title = "Select XML Input file for Raster Processor";
            dlg.Filter = "Raster Processor Input File(*.xml) | *.xml";

            DialogResult dr = dlg.ShowDialog();
            if (dr == DialogResult.Cancel)
            {
                return;
            }

            string[] selectedFiles;
            selectedFiles = dlg.FileNames;

            RunRasterProcessor(selectedFiles[0]);

        }

        public static void RunRasterProcessor(string inputFile)
        {
            //MessageBox.Show(string.Format("Executing {0}...", inputFile));
            //RasterAnalysis.ExecuteProcess(inputFile);
            //using (IRasterProgressTracker tracker = RasterProgressFactory.Create(ProgressTrackerCallBack))
            //{
            //  RasterAnalysis.ExecuteProcess(inputFile, tracker);
            //}
        }

        #region MapInfo_ProgressTracker_PRCOperation
        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>  CallBack method called by MIRaster library to provide progress feedback on UI. </summary>
        /// <param name="ppProcessProgress">   pointer to native structure  </param>
        ////////////////////////////////////////////////////////////////////////////////////////////////////
        private static void ProgressTrackerCallBack(IntPtr ppProcessProgress)
        {
            //Marshal the IntPtr to ProcessState type of struct
            ProcessState processProgress = (ProcessState)Marshal.PtrToStructure(ppProcessProgress, typeof(ProcessState));

            //Information can be used here.
            //Console.WriteLine("{0}", processProgress.Message);
            //Console.WriteLine("CurrentProgress: {0} %", processProgress.Progress);

            //If current operation needs to be paused, provide the cancel code as below.
            //RasterProgressFactory.PRCOperation(processProgress.TrackerHandle, RasterOperationProgress.Pause);

            //If current operation needs to be resumed, provide the cancel code as below.
            //RasterProgressFactory.PRCOperation(processProgress.TrackerHandle, RasterOperationProgress.Resume);

            //If current operation needs to be cancelled, provide the cancel code as below.
            //RasterProgressFactory.PRCOperation(processProgress.TrackerHandle, RasterOperationProgress.Cancel);
        }
        #endregion
    }
}
