import arcpy
import sys

if len(sys.argv) == 1:
    exit(1)

in_kml_file = sys.argv[1]
output_folder = sys.argv[2]

arcpy.KMLToLayer_conversion (in_kml_file, output_folder)

'''
arcpy.Buffer_analysis("roads", outfile, distance, "FULL", "ROUND", "LIST", "Distance")

field_area = 'Area'
minimum_area = 20
fcInput = 'ThePathAndOrNameOfMyFeatureClass'
fcOutput = 'ThePathAndOrNameOfTheOutputFeatureClass'
where_clause = '{} > {}'.format(arcpy.AddFieldDelimiters(fcInput, field_area), minimum_area)
arcpy.Select_analysis(fcInput, fcOutput, where_clause)
'''
