Compile AERMAP, AERMOD and AERPLOT Input and Run Output.
Live visualisation of polygon and point sources and automatic conversion (lon, lat to UTM) and input of coordinates via map.
It utilizes AERMAP (Terrain preprocessor of AERMOD), assumes the user has prepared elevation data (DEM or NED) and 
meteorology files (.src and .pfl), processes AERMAP output utilizing AERMOD (Steady-state plume model that 
incorporates air dispersion based on planetary boundary layer turbulence structure and scaling concepts, 
including treatment of both surface and elevated sources, and both simple and complex terrain, up to 50 km)
and finally visualizes maximall three averaging period utilizing AERPLOT (Postprocessor for the conversion of .plt files to .kml files).
config.json must be in same folder as CAIROforAERMOD.py, aermap.exe, aermod.exe, aerplot.exe are neccessary for running modeling.
Done as part of Masters Dissertation in Environmental Engineering at UNIVPM (Universita Politecnica delle Marche) 
by MSc Dominik Subotić 
under advisor Prof. Eng. Giorgio Passerini and curriculum supervisor PhD Simone Virgili.
