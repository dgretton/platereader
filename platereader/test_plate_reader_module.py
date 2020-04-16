from platereader.clariostar import ClarioStar

with ClarioStar() as reader_int:
    reader_int.plate_out()
    reader_int.run_protocols(['chemostat_od_abs', 'beta-gal'])
    reader_int.plate_out()
