using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using MapInfo.RasterEngine.Common;
using MapInfo.RasterEngine.Operations;
using MapInfo.RasterEngine.IO;

namespace MapInfoProAdvancedTool
{
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

            //Merge input grids with Sum operator
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
            
            //Merge input grids with Sum operator
            RasterProcessing.Merge(strInFilePaths, 0, strOutFilePath, rasterDriverID, MergeOperator.Average
                                , MergeType.Union, RasterResampleMethod.Bilinear, false
                                , MergeMultiResolutionMode.OptimumMaximum, null, apiOptions, null);

            return strOutFilePath;
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

    }
}
