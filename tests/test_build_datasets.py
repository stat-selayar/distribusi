import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import fitz

from build_datasets import (RAW_DIR, parse_id_number, parse_jampea_words_pdf,
                            parse_period_from_path)


class ParsePeriodFromPathTest(unittest.TestCase):
    def test_month_name_with_year_in_filename(self):
        path = 'Jampea/2026/LAP. OPS. UPP JAMPEA PERIODE MEI 2026.pdf'
        self.assertEqual(parse_period_from_path(path), (2026, 5))

    def test_numeric_month_prefix(self):
        for bulan in (1, 2, 4):
            path = f'Jampea/2026/2026_{bulan:02d}_730104_JAMPEA.pdf'
            self.assertEqual(parse_period_from_path(path), (2026, bulan))

    def test_year_falls_back_to_parent_directory(self):
        path = 'Jampea/2026/LAPORAN OPERASIONAL JUNI.pdf'
        self.assertEqual(parse_period_from_path(path), (2026, 6))

    def test_unrecognised_month_is_not_defaulted_to_january(self):
        path = 'Jampea/2026/LAPORAN OPERASIONAL.pdf'
        self.assertEqual(parse_period_from_path(path), (2026, None))


class ParseIdNumberTest(unittest.TestCase):
    def test_indonesian_and_plain_formats(self):
        self.assertEqual(parse_id_number('2.097'), 2097.0)
        self.assertEqual(parse_id_number('1.150'), 1150.0)
        self.assertEqual(parse_id_number('1,5'), 1.5)
        self.assertEqual(parse_id_number('22,00'), 22.0)
        self.assertEqual(parse_id_number('21.63'), 21.63)
        self.assertEqual(parse_id_number('842'), 842.0)

    def test_non_numbers_are_none(self):
        for s in ('-', '', None, '######', 'SEMEN'):
            self.assertIsNone(parse_id_number(s))


def make_jampea_pdf():
    """A one-page report with 36 vertical rules (35 columns) and text placed per column."""
    edges = [100 + 24 * i for i in range(36)]
    doc = fitz.open()
    page = doc.new_page(width=1000, height=300)
    for x in edges:
        page.draw_line((x, 40), (x, 260))

    def put(y, **cols):
        for col, text in cols.items():
            page.insert_text((edges[int(col[1:])] + 1, y), text, fontsize=4)

    put(100, c0='1', c1='KAPALA', c2='RI', c4='21,63', c5='22', c7='01-06-26', c9='BULUKUMBA',
        c12='KLM', c13='02-06-26', c15='SELAYAR', c16='LPG', c17='27', c20='BERAS', c21='1,5',
        c28='3', c29='4')
    put(103, c16='SEMEN', c17='16')
    put(106, c17='296', c21='152')  # unlabelled section subtotal
    put(109, c1='JUMLAH', c17='999')
    put(112, c0='#', c1='KAPALB', c16='-', c17='-')
    return doc


class JampeaWordsPdfTest(unittest.TestCase):
    def test_vessels_split_and_subtotals_ignored(self):
        recs = parse_jampea_words_pdf(make_jampea_pdf(), 2026, 6)
        self.assertEqual([r['nama_kapal'] for r in recs], ['KAPALA', 'KAPALB'])
        a = recs[0]
        self.assertEqual((a['tahun'], a['bulan'], a['tipe_kapal']), (2026, 6, 'KLM'))
        self.assertEqual((a['gt'], a['loa']), (22.0, 21.63))
        self.assertEqual((a['penumpang_turun'], a['penumpang_naik']), (3, 4))
        self.assertEqual((a['barang_bongkar_ton'], a['barang_muat_ton']), (43.0, 1.5))
        self.assertEqual(a['jenis_barang_bongkar'], ['LPG', 'SEMEN'])
        self.assertEqual((a['pelabuhan_asal'], a['pelabuhan_tujuan']), ('BULUKUMBA', 'SELAYAR'))

    def test_page_without_ruled_table_gives_no_records(self):
        doc = fitz.open()
        doc.new_page().insert_text((50, 50), 'LAPORAN', fontsize=10)
        self.assertEqual(parse_jampea_words_pdf(doc, 2026, 6), [])


JUNE_PDF = os.path.join(RAW_DIR, 'Jampea/2026/LAPORAN OPERASIONAL JUNI.pdf')


@unittest.skipUnless(os.path.exists(JUNE_PDF), 'raw data not available')
class JampeaJuneRealFileTest(unittest.TestCase):
    def test_first_vessel_and_count(self):
        recs = parse_jampea_words_pdf(fitz.open(JUNE_PDF), 2026, 6)
        self.assertEqual(len(recs), 163)
        first = recs[0]
        self.assertEqual((first['nama_kapal'], first['tanggal_tiba']), ('REZKY AQILA', '01-06-26'))
        self.assertEqual((first['barang_bongkar_ton'], first['barang_muat_ton']), (27.0, 17.0))


if __name__ == '__main__':
    unittest.main()
