#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# This script needs to be run with python above version 3.
# To install module xlutils, run the command: sudo pip install xlutils
# To install module python-pptx, run the command: sudp pip3 install python-pptx
"""
This script is a tool used to filter and analysis data from TSO500 results.
And generate the PP report based on the results data and template file.
"""

import os
import logging
import re
import sys
import shutil
import getopt
import time
from datetime import datetime
from configparser import ConfigParser
from xlutils.copy import copy
from pptx import Presentation
from pptx.util import Cm, Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_PARAGRAPH_ALIGNMENT
from pptx.enum.shapes import MSO_SHAPE
from decimal import Decimal
from copy import deepcopy
import pronto.pronto as pronto
import pandas
import math
from pdf2image import convert_from_path

runID = ""
DNA_sampleID = ""
RNA_sampleID = ""
extra_path = ""
batch_nr = ""
tumor_content_nr = ""
ipd_birth_year = ""
ipd_diagnosis_year = "-"
ipd_age = ""
ipd_gender = ""
ipd_consent = ""
ipd_collection_year = "-"
requisition_hospital = ""
pathology_comment = ""
ipd_material_id = ""
DNA_material_id = ""
RNA_material_id = ""
sample_info_comment = ""
extraction_hospital = ""
inclusion_site = ""
ipd_clinical_diagnosis = "-"
sample_material = ""
sample_type = ""
tumor_type = ""
TMB_DRUP = ""
str_TMB_DRUP = ""
TMB_TSO500 = ""
MSI_TSO500 = ""
pipline = ""

def read_tsv(data_file,filter_column,key_word):
	data = []
	mark = False
	col0 = "Sample_ID"
	global DNA_sampleID
	for line in open(data_file):
		line_cells = line.split('\t')
		if(line.split('\t')[0] == col0 and not mark):
			line_cells = line.split('\t')
			for col in range(len(line_cells)):
				if(line_cells[col] == "IGV_QC"):
					IGV_QC_col = col
				if(line_cells[col] == "Class_judgement"):
					Class_judgement_col = col
				if(line_cells[col] == filter_column):
					filter_column_n = col
			line_cells_string = [line_cells[i] + '\t' for i in range(len(line_cells))]
			line_cells_string.append('\n')
			data.append(line_cells_string)
			mark = True
		if(line.startswith(DNA_sampleID) and mark):
			if(line_cells[IGV_QC_col] == "Not OK" and line_cells[Class_judgement_col] != "exclude"):
				print ("""              Dataset error: 
			IGV_QC is 'Not OK', but Class_judement is not 'exclude'. Please check the QC Excel file and fix the mistake before run this script again!
				""")
				sys.exit(0)
			filter_column_data = line_cells[filter_column_n]
			if key_word.startswith('!'):
				if("&&" in key_word):
					not_key = key_word.replace('!','')
					key = not_key.split(" && ")
				else:
					key = key_word.split('!')
				appear = False
				for filter_column in filter_column_data.split(','):
					if(filter_column in key):
						appear = True
				if(appear == False):
					line_cells_string = [line_cells[i] + '\t' for i in range(len(line_cells))]
					line_cells_string.append('\n')
					data.append(line_cells_string)
			else:
				key = key_word.split(',')
				if(line_cells[filter_column_n] in key):
					line_cells_string = [line_cells[i] + '\t' for i in range(len(line_cells))]
					line_cells_string.append('\n')
					data.append(line_cells_string)
	return data


def read_tsv_col(data_file,filter_column,key_word,columns,MTB_format):
	mark = False
	col0 = "Sample_ID"
	global DNA_sampleID
	nlines = len(open(data_file).readlines())
	data = [[] for n in range(nlines)]
	column_mark = []
	columnNames = columns.split(',')
	d = 0
	# Verify the column numbers only for filtering the sequence_summary table. (MTB_format==True)
	RefSeq_mRNA_col = 3
	cDNA_change_col = 6
	Change_summary_col = 10
	Depth_tumor_DNA_col = 12
	AF_tumor_DNA_col = 13
	for line in open(data_file):
		line = line.replace('\n', '')
		line_cells = line.split('\t')
		if(line_cells[0] == col0 and not mark):
			for col in range(len(line_cells)):
				if(line_cells[col] == "IGV_QC"):
					IGV_QC_col = col
				if(line_cells[col] == "Class_judgement"):
					Class_judgement_col = col
				if(line_cells[col] == filter_column):
					filter_column_n = col
				if(line_cells[col] == "Coding_status"):
					Coding_status_col = col
				if(line_cells[col] == "Genomic_location"):
					Genomic_location_col = col
				if(line_cells[col] == "DNA_change"):
					DNA_change_col = col
				if(line_cells[col] == "Gene_symbol"):
					Gene_symbol_col = col
				for m in range(len(columnNames)):
					if(line_cells[col] in columnNames):
						line_cells_string = line_cells[col] + '\t'
						data[0].append(line_cells_string)
						column_mark.append(col)
						break
			if(MTB_format == True):
				data[0].insert(1,"Genomic_coordinates_in_hg19_build\t")
				data[0].insert(5,"HGVS_syntax\t")
				data[0].insert(6,"Change_summary\t")
				data[0].insert(8,"Read_depth(variant reads/total reads)\t")
			data[0].append('\n')
			mark = True
		if(line.startswith(DNA_sampleID) and mark):
			if(line_cells[IGV_QC_col] == "Not OK" and line_cells[Class_judgement_col] != "exclude"):
				print ("""              Dataset error: 
			IGV_QC is 'Not OK', but Class_judement is not 'exclude'. Please check the QC Excel file and fix the mistake before run this script again!
				 """)
				sys.exit(0)
			filter_column_data = line_cells[filter_column_n]
			if key_word.startswith('!'):
				if("&&" in key_word):
					not_key = key_word.replace('!','')
					key = not_key.split(" && ")
				else:
					key = key_word.split('!')
				appear = False
				for filter_column in filter_column_data.split(','):
					if(filter_column in key):
						appear = True
				if(appear == False):
					d += 1
					for num in column_mark:
						while(len(line_cells) <= num):
							line_cells.append('')
						if(num == Coding_status_col):
							line_cells[num] = line_cells[num].replace("_variant", "") + '\t'
						else:
							line_cells[num] = line_cells[num] + '\t'
						data[d].append(line_cells[num])
					if(MTB_format == True):
						try:
							MTB_format_str = "chr" + line_cells[Genomic_location_col].split(":")[0] + ":g." + line_cells[Genomic_location_col].split(":")[1].replace('\t','') + line_cells[DNA_change_col] + '\t'
							data[d].insert(1,MTB_format_str)
							HGVS_syntax_str = line_cells[RefSeq_mRNA_col] + ":" + line_cells[cDNA_change_col] + '\t'
							data[d].insert(5,HGVS_syntax_str)
							include_exon = True
							Change_summary_format = ""
							Change_summary_ori = line_cells[Change_summary_col].replace(":NA", "")
							for string in Change_summary_ori.split(":"):
								if("exon" in string):
									include_exon = False
									continue
								else:
									if(include_exon):
										Change_summary_format += string
									else:
										Change_summary_format += ",p.(" + string + ")"
							Change_summary_str = line_cells[Gene_symbol_col].replace('\t',' ') + line_cells[RefSeq_mRNA_col] + ":" + Change_summary_format + '\t'
							data[d].insert(6,Change_summary_str)
							variant_reads_str = int(int(line_cells[Depth_tumor_DNA_col]) * float(line_cells[AF_tumor_DNA_col]))
							total_reads_str = str(line_cells[Depth_tumor_DNA_col])
							Read_depth_str = str(variant_reads_str) + "/" + total_reads_str + '\t'
							data[d].insert(8,Read_depth_str)
							AF_tumor_DNA_ori = float(line_cells[AF_tumor_DNA_col]) * 100
							AF_tumor_DNA_str = format(AF_tumor_DNA_ori, '.1f') + '%'
							data[d][9] = AF_tumor_DNA_str  + '\t'
						except:
							print("Warning: Data issue for " + DNA_sampleID + "! Please check the small_variant_table from TSOPPI. The report will be generated for further QC.")
			else:
				key = key_word.split(',')
				if(line_cells[filter_column_n] in key):
					d += 1
					for num in column_mark:
						if(num == Coding_status_col):
							line_cells[num] = line_cells[num].replace("_variant", "") + '\t'
						else:
							line_cells[num] = line_cells[num] + '\t'
						data[d].append(line_cells[num])
			data[d].append('\n')
	return data


def filter_depth_tumor_cols(data_config,depth_tumor_DNA):
	data = [[] for n in range(len(data_config))]
	data[0] = data_config[0]
	p = data_config[0].index('Depth_tumor_DNA\t')
	for i in range(len(data_config)):
		for j in range(len(data_config[i])):
			if(data_config[i][p] != 'Depth_tumor_DNA\t' and data_config[i][p] != ''):
				num = data_config[i][p].split('\t')[0]
				if(int(num) >= depth_tumor_DNA):
					data[i].append(data_config[i][j])
	return data


def write_exl(output_file,data):
	file_dir = os.path.split(output_file)[0]
	if not os.path.exists(file_dir):
		os.makedirs(file_dir)
	txt_file = open(output_file, mode='w', encoding='utf-8')
	for item in data:
		data_string = str(item) + "\t"
		txt_file.writelines(item)
	txt_file.close()


def clear_blank_line(file_in,file_out):
	fr = open(file_in, 'r')
	fw = open(file_out, 'w')
	for line in fr.readlines():
		if(line.split()):
			fw.write(line)
	fr.close()
	fw.close()
	os.remove(file_in)


def get_patient_info_from_MTF_2023(ipd_material_file,ipd_no,DNA_sampleID,RNA_sampleID):
	import xlrd
	global ipd_birth_year
	global ipd_gender
	global ipd_consent
	global DNA_material_id
	global RNA_material_id
	global ipd_collection_year
	global requisition_hospital
	global extraction_hospital
	global batch_nr
	global tumor_content_nr
	global inclusion_site
	open_exl_material = xlrd.open_workbook(ipd_material_file)
	sheet_material = open_exl_material.sheet_by_index(0)
	nrows_material = sheet_material.nrows
	ncols_material = sheet_material.ncols
	columns = {
		'ipd': 'InPreD ID',
		'gender': 'Gender',
		'age': 'Age',
		'birth_date': 'Date of birth',
		'requisition_hospital': 'Requester Hospital',
		'material_name': 'Original Name',
		'consent': 'Study ID',
		'tumor_content_nr': 'Tumor cells [%]',
		'sampleID_ori_name': 'Sample ID',
		'ex_sample_info': 'Sample information',
		'ex_data_section': 'Extraction Data',
		'ex_library_pre': 'Library Preparation (LP) Data',
		'extraction_hospital': 'Extraction Hospital',
		'batch_nr': 'LP batch'
	}
	ipd_birth_date = ""
	sample_info_row = 0
	extra_data_row = 0
	library_pre_row = 0
	for l in range(nrows_material):
		if(sheet_material.cell_value(l,0) == columns['ex_sample_info']):
			sample_info_row = l
		if(sheet_material.cell_value(l,0) == columns['ex_data_section']):
			extra_data_row = l
		if(sheet_material.cell_value(l,0) == columns['ex_library_pre']):
			library_pre_row = l
	for r in range(nrows_material):
		for c in range(ncols_material):
			if(sheet_material.cell_value(r,c) == columns['ipd']):
				ipd_MTF = sheet_material.cell_value(r+2,c)
				if(ipd_MTF != ipd_no):
					print("""               Error:
			The InPreD patient ID in IPD Material Transit Form InPreD NGS file does not match with the IPD number! 
			Please check and fix the mistake before run this script again!""")
					print("                 IPD is " + ipd_MTF + " in MTF, while IPD is " + ipd_no[3:] + " in TSO500.")
					sys.exit(0)
			if(sheet_material.cell_value(r,c) == columns['birth_date']):
				ipd_birth_date_exl = sheet_material.cell_value(r+2,c)
				try:
					datetime_date = str(xlrd.xldate_as_datetime(ipd_birth_date_exl,0))
					ipd_birth_year = datetime_date.split('-')[0]
				except:
					ipd_birth_year = "-"
			if(sheet_material.cell_value(r,c) == columns['gender'] and ipd_gender == ""):
				gender = str(sheet_material.cell_value(r+2,c))
				if(gender != "" and gender != "X"):
					ipd_gender = gender[0]
			if(sheet_material.cell_value(r,c) == columns['age']):
				ipd_age = str(sheet_material.cell_value(r+2,c))
			if(sheet_material.cell_value(r,c) == columns['consent'] and ipd_consent == ""):
				ipd_consent = str(sheet_material.cell_value(r+2,c))
				for r in range(r,(sample_info_row-2)):
					if(ipd_consent == "0.0"):
						ipd_consent = ""
					if(sheet_material.cell_value(r,6) == columns['requisition_hospital'] and requisition_hospital == ""):
						requisition_hospital = sheet_material.cell_value(r+2,6)
					if((sheet_material.cell_value(r+2,c) != "" or sheet_material.cell_value(r+2,c) != "-" or sheet_material.cell_value(r+2,c) != "0.0") and str(sheet_material.cell_value(r+2,c)) not in ipd_consent):
						if(ipd_consent == ""):
							ipd_consent = str(sheet_material.cell_value(r+2,c))
						else:
							ipd_consent = ipd_consent + "," + str(sheet_material.cell_value(r+2,c))
						continue
			if(sheet_material.cell_value(r,c) == columns['material_name'] and ipd_material_id == ""):
				for r in range(r,(extra_data_row-2)):
					if(sheet_material.cell_value(r+2,9) == DNA_sampleID and sheet_material.cell_value(r+2,c) != "" and str(sheet_material.cell_value(r+2,c)) not in DNA_material_id):
						if(DNA_material_id == ""):
							DNA_material_id = str(sheet_material.cell_value(r+2,c))
						else:
							DNA_material_id = DNA_material_id + "," + str(sheet_material.cell_value(r+2,c))
						tumor_content_nr = sheet_material.cell_value(r+2,2)
						continue
					if(RNA_sampleID != "" and sheet_material.cell_value(r+2,9) == RNA_sampleID and sheet_material.cell_value(r+2,c) != "" and str(sheet_material.cell_value(r+2,c)) not in RNA_material_id):
						if(RNA_material_id == ""):
							RNA_material_id = str(sheet_material.cell_value(r+2,c))
						else:
							RNA_material_id = RNA_material_id + "," + str(sheet_material.cell_value(r+2,c))
						continue
			if(sheet_material.cell_value(r,c) == columns['extraction_hospital'] and extraction_hospital == ""):
				for r in range(r,(library_pre_row-2)):
					if(sheet_material.cell_value(r+2,8) == DNA_sampleID):
						extraction_hospital = str(sheet_material.cell_value(r+2,c))
						break
			if(sheet_material.cell_value(r,c) == columns['batch_nr'] and batch_nr == ""):
				for r in range(r,(nrows_material-2)):
					if(sheet_material.cell_value(r+2,0) == DNA_sampleID):
						batch_nr = str(sheet_material.cell_value(r+2,c))
	open_exl_material.release_resources()
	if(ipd_consent == "0.0"):
		ipd_consent = ""
	if(ipd_age == "" and ipd_birth_date != ""):
		ipd_age = "<1"
	inclusion_site_list = {'R': 'Radium', 'U': 'Ullevål', 'C': 'Riksen', 'A': 'Ahus', 'D': 'Drammen', 'B': 'Bærum', 'G': 'Gjøvik', 'I': 'Hamar', 'L': 'Lillehammer', 'T': 'Vestfold', 'K': 'Sørlandet', 'Q': 'Østfold', 'V': 'Telemark', 'Y': 'Lovisenberg', 'H': 'Haukeland', 'S': 'Stavanger', 'E': 'Fonna', 'F': 'Førde', 'O': 'St.Olavs', 'M': 'Nord-trøndelag', 'J': 'Møre og Romsdal', 'N': 'Nord Norge', 'P': 'Nordland'}
	if("IKKE IMPRESS" in ipd_consent):
		inclusion_site = ""
	else:
		try:
			site_letter_code = ipd_consent[-6]
			inclusion_site = inclusion_site_list.get(site_letter_code)
		except:
			inclusion_site = "Inclusion site"


# Read clinical data from MTF files version from 2024 and newer.
def get_patient_info_from_MTF_new(ipd_material_file,ipd_no,DNA_sampleID,RNA_sampleID):
	import xlrd
	global ipd_birth_year
	global ipd_age
	global ipd_clinical_diagnosis
	global ipd_gender
	global ipd_consent
	global DNA_material_id
	global RNA_material_id
	global ipd_collection_year
	global requisition_hospital
	global pathology_comment
	global sample_info_comment
	global extraction_hospital
	global batch_nr
	global tumor_content_nr
	global inclusion_site
	open_exl_material = xlrd.open_workbook(ipd_material_file)
	sheet_material = open_exl_material.sheet_by_index(0)
	nrows_material = sheet_material.nrows
	ncols_material = sheet_material.ncols
	columns = {
		'ipd': 'InPreD ID',
		'gender': 'Gender',
		'age': 'Age',
		'birth_date': 'Date of birth',
		'DIT_number': 'DIT number',
		'consent': 'Study ID',
		'requisition_hospital': 'Requester Hospital',
		'Histopathological_diagnosis': 'diagnosis',
		'comment': 'Comments',
		'material_id': 'Sample material ID',
		'tumor_content_nr': 'Tumor cells [%]',
		'sample_ID': 'Sample ID',
		'ex_pathology_info': 'Molecular Pathology information',
		'ex_sample_info': 'Sample information',
		'ex_data_section': 'Extraction Data',
		'ex_library_pre': 'Library Preparation (LP) Data',
		'extraction_hospital': 'Extraction Hospital',
		'batch_nr': 'LP batch'
		}
	ipd_birth_date = ""
	sample_info_row = 0
	extra_data_row = 0
	library_pre_row = 0
	for l in range(nrows_material):
		if(sheet_material.cell_value(l,0) == columns['ex_sample_info']):
			sample_info_row = l
		if(sheet_material.cell_value(l,0) == columns['ex_data_section']):
			extra_data_row = l
		if(sheet_material.cell_value(l,0) == columns['ex_library_pre']):
			library_pre_row = l  
	for r in range(nrows_material):
		for c in range(ncols_material):
			if(sheet_material.cell_value(r,c) == columns['ipd']):
				ipd_MTF = sheet_material.cell_value(r+2,c)
				if(ipd_MTF != ipd_no):
					print("""               Error:
                        The InPreD patient ID in IPD Material Transit Form InPreD NGS file does not match with the IPD number! 
                        Please check and fix the mistake before run this script again!""")
					print("                 IPD is " + ipd_MTF + " in MTF, while IPD is " + ipd_no[3:] + " in TSO500.")
					sys.exit(0)
			if(sheet_material.cell_value(r,c) == columns['birth_date']):
				ipd_birth_date_exl = sheet_material.cell_value(r+2,c)
				try:
					datetime_date = str(xlrd.xldate_as_datetime(ipd_birth_date_exl,0))
					ipd_birth_year = datetime_date.split('-')[0]
				except:
					try:
						day, month, year = map(int, ipd_birth_date_exl.split('.'))
						datetime(year,month,day)
						ipd_birth_year = str(year)
					except:
						ipd_birth_year = "-"
			if(sheet_material.cell_value(r,c) == columns['gender'] and ipd_gender == ""):
				gender = str(sheet_material.cell_value(r+2,c))
				if(gender != "" and gender != "X"):
					ipd_gender = gender[0]
			if(sheet_material.cell_value(r,c) == columns['age']):
				ipd_age = str(sheet_material.cell_value(r+2,c))
			if(sheet_material.cell_value(r,c) == columns['Histopathological_diagnosis'] and sheet_material.cell_value(r+1,c) != ""):
				ipd_clinical_diagnosis = str(sheet_material.cell_value(r+1,c))
			if(sheet_material.cell_value(r,c) == columns['consent'] and ipd_consent == ""):
				ipd_consent = str(sheet_material.cell_value(r+2,c))
				if(ipd_consent == "0.0"):
					ipd_consent = ""
				for r in range(r,(sample_info_row-2)):
					DIT_number = "-"
					comments = "-"
					if(sheet_material.cell_value(r,0) == columns['DIT_number'] and sheet_material.cell_value(r+2,0) != ""):
						DIT_number = sheet_material.cell_value(r+2,0)
					if(sheet_material.cell_value(r,6) == columns['requisition_hospital'] and requisition_hospital == ""):
						requisition_hospital = sheet_material.cell_value(r+2,6)
					if((sheet_material.cell_value(r+2,c) != "" and sheet_material.cell_value(r+2,c) != "-" and str(sheet_material.cell_value(r+2,c)) != "0.0") and str(sheet_material.cell_value(r+2,c)) not in ipd_consent):
						if(ipd_consent == ""):
							ipd_consent = str(sheet_material.cell_value(r+2,c))
						else:
							ipd_consent = ipd_consent + "," + str(sheet_material.cell_value(r+2,c))
					if(sheet_material.cell_value(r,10) == columns['comment'] and sheet_material.cell_value(r+2,10) != "" and str(sheet_material.cell_value(r+2,10)) != "0.0" and str(sheet_material.cell_value(r+2,10)) != "0"):
						comments = str(sheet_material.cell_value(r+2,10)).replace("\n", " ")
					if(pathology_comment == ""):
						pathology_comment = DIT_number + ":" + comments
					else:
						if(DIT_number != "-" or comments != "-"):
							pathology_comment += "|" + DIT_number + ":" + comments
			if(sheet_material.cell_value(r,c) == columns['material_id'] and ipd_material_id == ""):
				for r in range(r,(extra_data_row-2)):
					sample_ID = "-"
					comments = "-"
					if(sheet_material.cell_value(r+2,9) == DNA_sampleID and sheet_material.cell_value(r+2,c) != "" and str(sheet_material.cell_value(r+2,c)) not in DNA_material_id):
						if(DNA_material_id == ""):
							DNA_material_id = str(sheet_material.cell_value(r+2,c))
						else:
							DNA_material_id = DNA_material_id + "," + str(sheet_material.cell_value(r+2,c))
						tumor_content_nr = sheet_material.cell_value(r+2,2)
					if(RNA_sampleID != "" and sheet_material.cell_value(r+2,9) == RNA_sampleID and sheet_material.cell_value(r+2,c) != "" and str(sheet_material.cell_value(r+2,c)) not in RNA_material_id):
						if(RNA_material_id == ""):
							RNA_material_id = str(sheet_material.cell_value(r+2,c))
						else:
							RNA_material_id = RNA_material_id + "," + str(sheet_material.cell_value(r+2,c))
					if(sheet_material.cell_value(r+2,9) != ""):
						sample_ID = sheet_material.cell_value(r+2,9)
					if(sheet_material.cell_value(r+2,10) != "" and str(sheet_material.cell_value(r+2,10)) != "0.0" and str(sheet_material.cell_value(r+2,10)) != "0"):
						comments = str(sheet_material.cell_value(r+2,10)).replace("\n", " ")
					if(sample_info_comment == ""):
						sample_info_comment = "{}: {}".format(sample_ID, comments)
					else:
						if(sample_ID != "-" or comments != "-"):
							sample_info_comment += "|" + sample_ID + ": " + comments
			if(sheet_material.cell_value(r,c) == columns['extraction_hospital'] and extraction_hospital == ""):
				for r in range(r,(library_pre_row-2)):
					if(sheet_material.cell_value(r+2,9) == DNA_sampleID):
						try:
							extraction_hospital = str(sheet_material.cell_value(r+2,c)).split(",")[1]
						except:
							extraction_hospital = str(sheet_material.cell_value(r+2,c))
						break
			if(sheet_material.cell_value(r,c) == columns['batch_nr'] and batch_nr == ""):
				for r in range(r,(nrows_material-2)):
					if(sheet_material.cell_value(r+2,0) == DNA_sampleID):
						batch_nr = str(sheet_material.cell_value(r+2,c))
	open_exl_material.release_resources()
	if(ipd_consent == "0.0"):
		ipd_consent = ""
	if(ipd_age == "" and ipd_birth_date != ""):
		ipd_age = "<1"
	inclusion_site_list = {'R': 'Radium', 'U': 'Ullevål', 'C': 'Riksen', 'A': 'Ahus', 'D': 'Drammen', 'B': 'Bærum', 'G': 'Gjøvik', 'I': 'Hamar', 'L': 'Lillehammer', 'T': 'Vestfold', 'K': 'Sørlandet', 'Q': 'Østfold', 'V': 'Telemark', 'Y': 'Lovisenberg', 'H': 'Haukeland', 'S': 'Stavanger', 'E': 'Fonna', 'F': 'Førde', 'O': 'St.Olavs', 'M': 'Nord-trøndelag', 'J': 'Møre og Romsdal', 'N': 'Nord Norge', 'P': 'Nordland'}
	if("IKKE IMPRESS" in ipd_consent):
		inclusion_site = ""
	else:
		try:
			site_letter_code = ipd_consent[-6]
			inclusion_site = inclusion_site_list.get(site_letter_code)
		except:
			inclusion_site = "Inclusion site"


def get_RNA_material_id(InPreD_clinical_data_file,RNA_sampleID,encoding_sys):
	RNA_material_id_exist = False
	if(encoding_sys != ""):
		f = open(InPreD_clinical_data_file, 'r', encoding=encoding_sys)
	else:
		f = open(InPreD_clinical_data_file, 'r')
	for l in f:
		if(RNA_sampleID == l.split('\t')[0]):
			RNA_material_id = l.split('\t')[8]
			RNA_material_id_exist = True
	f.close()
	if(RNA_material_id_exist == False):
		print("Warning: The "+ RNA_sampleID + " does not exist in the meta file! The report will be generated without RNA sample material ID!")
		RNA_material_id = ""
	return RNA_material_id


def update_ppt_template_data(inpred_node,ipd_no,ipd_gender,ipd_age,ipd_diagnosis_year,DNA_material_id,RNA_material_id,ipd_consent,requisition_hospital,pathology_comment,ipd_clinical_diagnosis,tumor_type,sample_type,sample_material,sample_info_comment,pipline,tumor_content,ppt_template,output_ppt_file):
	if(ipd_age != "" and ipd_age != "-" and ipd_age != "XX" and ipd_age != "<1"):
		age = str(int(float(ipd_age)))
	else:
		age = ipd_age
	try:
		sample = sample_type + '\n' + sample_material
	except:
		sample = ""
	today_date = time.strftime("%d", time.localtime())
	today_month = time.strftime("%b", time.localtime())
	today_year = time.strftime("%Y", time.localtime())
	today = today_date + '\n' + today_month.upper() + '\n' + today_year
	year_text =  "XX\nXXX\n" + today_year
	ppt = Presentation(ppt_template)
	indexs = [1,3,4,5,6]
	for index in indexs:
		slide = ppt.slides[index]
		textbox0 = slide.shapes.add_textbox(Inches(0.06), Inches(0.06), Inches(0.47), Inches(0.63))
		tf0 = textbox0.text_frame
		tf0.paragraphs[0].text = year_text
		tf0.paragraphs[0].font.size = Pt(11)
		tf0.paragraphs[0].font.color.rgb = RGBColor(250,250,250)
		tf0.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox1 = slide.shapes.add_textbox(Inches(3.75), Inches(0.11), Inches(1.33), Inches(0.50))
		tf1 = textbox1.text_frame
		tf1.paragraphs[0].text = ipd_no
		tf1.paragraphs[0].font.size = Pt(24)
		tf1.paragraphs[0].font.color.rgb = RGBColor(250,250,250)
		tf1.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox2 = slide.shapes.add_textbox(Inches(8.99), Inches(0.02), Inches(0.45), Inches(0.55))
		tf2 = textbox2.text_frame
		tf2.paragraphs[0].text = today
		tf2.paragraphs[0].font.size = Pt(9)
		tf2.paragraphs[0].font.color.rgb = RGBColor(250,250,250)
		tf2.paragraphs[0].alignment = PP_ALIGN.CENTER
		tf2.vertical_anchor = MSO_VERTICAL_ANCHOR.BOTTOM
		textbox3 = slide.shapes.add_textbox(Inches(7.23), Inches(0.52), Inches(2.53), Inches(0.21))
		tf3 = textbox3.text_frame
		tf3.paragraphs[0].text = pipline
		tf3.paragraphs[0].font.size = Pt(7)
		tf3.paragraphs[0].font.color.rgb = RGBColor(64,64,64)
		textbox4 = slide.shapes.add_textbox(Inches(0.50), Inches(1.47), Inches(0.86), Inches(0.25))
		tf4 = textbox4.text_frame
		tf4.paragraphs[0].text = requisition_hospital
		tf4.paragraphs[0].font.size = Pt(10)
		tf4.paragraphs[0].alignment = PP_ALIGN.CENTER
		tf4.paragraphs[0].font.color.rgb = RGBColor(250,250,250)
		textbox5 = slide.shapes.add_textbox(Inches(0.71), Inches(1.84), Inches(0.86), Inches(0.50))
		tf5 = textbox5.text_frame
		tf5.paragraphs[0].text = sample
		tf5.paragraphs[0].font.size = Pt(8)
		tf5.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox6 = slide.shapes.add_textbox(Inches(0.81), Inches(2.65), Inches(0.63), Inches(0.33))
		tf6 = textbox6.text_frame
		tf6.paragraphs[0].text = tumor_content
		tf6.paragraphs[0].font.size = Pt(14)
		tf6.paragraphs[0].alignment = PP_ALIGN.CENTER
		if(index == 1 or ipd_clinical_diagnosis == "-" or ipd_clinical_diagnosis == ""):
			textbox7 = slide.shapes.add_textbox(Inches(5.77), Inches(0.19), Inches(0.86), Inches(0.33))
			tf7 = textbox7.text_frame
			tf7.paragraphs[0].text = str(tumor_type)
			tf7.paragraphs[0].font.size = Pt(14)
		else:
			if('\n' in ipd_clinical_diagnosis):
				textbox7 = slide.shapes.add_textbox(Inches(4.95), Inches(0.11), Inches(2.36), Inches(0.50))
				tf7 = textbox7.text_frame
				tf7.paragraphs[0].font.size = Pt(12)
			else:
				textbox7 = slide.shapes.add_textbox(Inches(5.77), Inches(0.19), Inches(0.86), Inches(0.33))
				tf7 = textbox7.text_frame
				tf7.paragraphs[0].font.size = Pt(14)
			tf7.paragraphs[0].text = ipd_clinical_diagnosis
		tf7.paragraphs[0].font.italic = True
		tf7.paragraphs[0].font.color.rgb = RGBColor(250,250,250)
		tf7.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox11 = slide.shapes.add_textbox(Inches(0.85), Inches(1.12), Inches(0.48), Inches(0.27))
		tf11 = textbox11.text_frame
		tf11.paragraphs[0].text = ipd_diagnosis_year
		tf11.paragraphs[0].font.size = Pt(10)
		tf11.paragraphs[0].alignment = PP_ALIGN.LEFT
		textbox12 = slide.shapes.add_textbox(Inches(0.61), Inches(0.35), Inches(1.02), Inches(0.33))
		tf12 = textbox12.text_frame
		tf12.paragraphs[0].text = inpred_node
		tf12.paragraphs[0].font.size = Pt(14)
		tf12.paragraphs[0].alignment = PP_ALIGN.CENTER
		tf12.paragraphs[0].font.color.rgb = RGBColor(250,250,250)
		if(index == 1):
			gender_age = ""
			ipd_material_id_index = ""
			ipd_consent_index = ""
		textbox8 = slide.shapes.add_textbox(Inches(0.69), Inches(0.79), Inches(0.87), Inches(0.40))
		tf8 = textbox8.text_frame
		tf8.paragraphs[0].text = gender_age
		tf8.paragraphs[0].font.size = Pt(18)
		tf8.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox9 = slide.shapes.add_textbox(Inches(0.73), Inches(2.25), Inches(0.70), Inches(0.26))
		tf9 = textbox9.text_frame
		tf9.paragraphs[0].text = ipd_material_id_index
		tf9.paragraphs[0].font.size = Pt(5)
		tf9.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox10 = slide.shapes.add_textbox(Inches(2.10), Inches(0.11), Inches(1.07), Inches(0.50))
		tf10 = textbox10.text_frame
		tf10.paragraphs[0].text = ipd_consent_index
		tf10.paragraphs[0].font.size = Pt(14)
		tf10.paragraphs[0].alignment = PP_ALIGN.CENTER
		tf10.paragraphs[0].font.italic = True
		tf10.paragraphs[0].font.color.rgb = RGBColor(250,250,250)
		if(index == 3):
			textbox11 = slide.shapes.add_textbox(Inches(1.85), Inches(1.25), Inches(3.25), Inches(0.27))
			tf11 = textbox11.text_frame
			tf11.paragraphs[0].text = pathology_comment + "\n\n" + sample_info_comment.replace("|","\n")
			tf11.paragraphs[0].font.size = Pt(10)
			tf11.paragraphs[0].alignment = PP_ALIGN.LEFT
		gender_age = "{}/{}y".format(ipd_gender, age)
		if(RNA_material_id != ""):
			ipd_material_id_index = "DNA:" + DNA_material_id + "\nRNA:" + RNA_material_id
		else:
			ipd_material_id_index = "DNA:" + DNA_material_id
		ipd_consent_index = "Trial ID\n" + ipd_consent 

	ppt.save(output_ppt_file)


def insert_image_to_ppt(DNA_sampleID,DNA_normal_sampleID,RNA_sampleID,DNA_image_path,RNA_image_path,output_ppt_file):
	DNA_image = []
	RNA_image = []
	image_mark = "sample_QC_plot.png"
	ppt = Presentation(output_ppt_file)
	slide = ppt.slides[4]
	shapes = slide.shapes

	for file in os.listdir(DNA_image_path):
		if(DNA_sampleID in file and image_mark in file):
			image = os.path.join(DNA_image_path,file)
			DNA_image.append(image)
		if(DNA_normal_sampleID != "" and DNA_normal_sampleID in file and image_mark in file):
			image = os.path.join(DNA_image_path,file)
			DNA_image.append(image)

	if(RNA_image_path != ""):
		for file in os.listdir(RNA_image_path):
			if(RNA_sampleID in file and image_mark in file):
				image = os.path.join(RNA_image_path,file)
				RNA_image.append(image)
	left0 =  Inches(3.20)
	width0 = Inches(3.30)
	height0 = Inches(1.75)
	if(DNA_image != ''):
		top = Inches(1.55)
		d = 0
		for images in DNA_image:
			left = left0 + d * width0
			pic = slide.shapes.add_picture(images,left,top,width0,height0)
			d = d + 1	
	if(RNA_image != ''):
		top = Inches(3.44)
		r = 0
		for images in RNA_image:
			left = left0 + r * width0
			pic = slide.shapes.add_picture(images,left,top,width0,height0)
			r = r + 1
	ppt.save(output_ppt_file)


def insert_table_to_ppt(table_file,slide_n,table_name,left_h,top_h,width_h,left_t,top_t,width_t,height_t,font_size,table_header,output_ppt_file,print_row_num,table_column_width,table_max_rows_per_slide):

	# load table data
	try:
		table_data = pandas.read_csv(table_file, sep='\t', keep_default_na=False)
	except pandas.errors.EmptyDataError:
		logging.warning("{} is empty".format(table_file))
		return
	
	# add empty columns for missing header columns and move additional columns to the right
	table_data = pronto.normalize_column_index(table_data, table_header)

	# round floats to 2 decimal places
	table_data = pronto.set_column_to_2_decimals(table_data, "AF_tumor_DNA")

	# determine column and row number
	cols = len(table_header)
	rows = len(table_data)

	# how many slides are required
	if not table_max_rows_per_slide:
		table_max_rows_per_slide = rows
	total_slides_needed = math.ceil(rows / table_max_rows_per_slide)

	# Add data to ppt
	ppt = Presentation(output_ppt_file)
	total_slides = len(ppt.slides)
	for slide_idx in range(total_slides_needed):
		current_slide_data = pronto.get_slide_table_data(table_data, slide_idx, table_max_rows_per_slide)
		if(total_slides_needed == 1 and slide_n <= total_slides):
			shapes = ppt.slides[slide_n - 1].shapes
		else:
			shapes = ppt.slides.add_slide(ppt.slide_layouts[6]).shapes

		# create new table on slide
		left = Inches(left_t)
		top = Inches(top_t)
		width = Inches(width_t)
		height = Inches(height_t)
		table_rows = len(current_slide_data)
		table = shapes.add_table(table_rows,cols,left,top,width,height).table

		# if table_column_width is provided, set the column width
		if len(table_column_width) == cols:
			for col_idx, width in enumerate(table_column_width):
				table.columns[col_idx].width = Inches(width)
		
		# fill in the table data and set font size
		for row_idx, row in enumerate(table.rows):
			for col_idx, cell in enumerate(row.cells):
				cell.text = str(current_slide_data[row_idx][col_idx])
				cell.text_frame.paragraphs[0].font.size = Pt(font_size)

		# add table title
		pronto.add_table_name(shapes, table_name, left_h, top_h, width_h, 0.25, 8, print_row_num, slide_idx, total_slides_needed, rows)

	ppt.save(output_ppt_file)
	return rows


def update_ppt_variant_summary_table(data_nrows,DNA_sampleID,RNA_sampleID,TMB_DRUP_nr,TMB_DRUP_str,DNA_variant_summary_file,RNA_variant_summary_file,output_file_preMTB_AppendixTable,output_table_file_filterResults_AllReporVariants_CodingRegion,output_ppt_file):
	DNA_summary_file = open(DNA_variant_summary_file)
	global str_TMB_DRUP
	global TMB_TSO500
	global MSI_TSO500
	for line in DNA_summary_file:
		if(line.startswith(DNA_sampleID)):
			if(line.split('\t')[1] == 'NA'):
				TMB_illumina = "TMB = NA"
			else:
				TMB_illumina = "TMB = " + line.split('\t')[1]
				TMB_TSO500 = line.split('\t')[1]
				TMB_TSO500_nr = round(float(TMB_TSO500.split()[0]))
			if(line.split('\t')[2] == 'NA'):
				MSI_illumina = "MSI = NA"
				msi_text = "-"
				stable_text = "NA"
				msi_stable = "Not available"
			else:
				MSI_illumina = "MSI = " + line.split('\t')[2]
				MSI_TSO500 = line.split('\t')[2]
			# X(Y/Z) 
			# Z<40: evaluation not reliable. Z>=40 && X=<10: MSI/Stable. Z>=40 && X>=20: MSI/Unstable. Z>=40 && 10<X<20: MSI/Likely unstable.
				X = float(MSI_TSO500.split()[0])
				YZ = MSI_TSO500.split("(")[1].split(")")[0]
				Y = int(YZ.split('/')[0])
				Z = int(YZ.split('/')[1])
				if(Z >= 40 and X >= 20):
					msi_text = "MS"
					stable_text = "Unstable"
				if(Z >= 40 and X <= 10):
					msi_text = "MS"
					stable_text = "Stable"
				if(Z >= 40 and 10 < X < 20):
					msi_text = "MS"
					stable_text = "Likely unstable"
				if(Z < 40):
					msi_text = "MS"
					stable_text = "Not reliable"
				msi_stable = str(Y) + " unstable out of " + str(Z)
			stable_text_long = stable_text + '\n' + msi_stable

	DNA_summary_file.close()
	
	splicing = "splicing: Not Assayed"
	fusion = "fusion: Not Assayed"	
	if(RNA_sampleID != ""):
		RNA_summary_file = open(RNA_variant_summary_file)
		for line in RNA_summary_file:
			if(line.startswith(RNA_sampleID)):
				splice_variants_str = line.split('\t')[4]
				if(splice_variants_str == 'NA'):
					splicing = "splicing: None reported"
				else:
					splicing = "splicing: " + splice_variants_str.split('(')[0]
					splicing_end = re.findall('\\|(.*?)\\(', splice_variants_str)
					if(splicing_end):
						for splice in splicing_end:
							splicing += ',' + splice
				gene_fusion_str = line.split('\t')[5]
				if(gene_fusion_str == 'NA\n'):
					fusion = "fusion: None reported"
				else:
					fusion = "fusion: " + gene_fusion_str.split('(')[0]
					fusion_end = re.findall('\\|(.*?)\\(', gene_fusion_str)
					if(fusion_end):
						for fus in fusion_end:
							fusion += ';' + fus
		RNA_summary_file.close()

	table_file_coding_region = open(output_table_file_filterResults_AllReporVariants_CodingRegion)
	appendix_nrows = len(table_file_coding_region.readlines()) - 1
	if(appendix_nrows == -1):
		appendix_nrows = "NA"
	table_file_preMTBTable_Appendix = open(output_file_preMTB_AppendixTable)
	preMTB_appendix_nrows = len(table_file_preMTBTable_Appendix.readlines()) - 1
	if(preMTB_appendix_nrows == -1):
		preMTB_appendix_nrows = "NA"
	if(TMB_DRUP_str == "-1"):
		str_TMB_DRUP = "NA"
		TMB_DRUP_str = "NA"
	else:
		effect_panel_size = float(TMB_DRUP_str.split('/')[1])
		if(effect_panel_size < 1.14):
			str_TMB_DRUP = "NA"
		else:
			str_TMB_DRUP = str(TMB_DRUP_nr)

	ppt = Presentation(output_ppt_file)
	indexs = [1,5,6]
	for index in indexs:
		slide = ppt.slides[index]
		shapes = slide.shapes
		textbox1 = slide.shapes.add_textbox(Inches(5.76), Inches(1.60), Inches(0.41), Inches(0.27))
		tf1 = textbox1.text_frame
		tf1.paragraphs[0].text = msi_text
		tf1.paragraphs[0].font.size = Pt(10)
		tf1.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox2 = slide.shapes.add_textbox(Cm(15.84), Cm(4.18), Cm(2.26), Cm(0.64))
		tf2 = textbox2.text_frame
		tf2.paragraphs[0].text = stable_text_long
		tf2.paragraphs[0].font.size = Pt(7)
		tf2.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox3 = slide.shapes.add_textbox(Inches(3.66), Inches(1.27), Inches(0.30), Inches(0.22))
		tf3 = textbox3.text_frame
		tf3.paragraphs[0].text = str(data_nrows)
		tf3.paragraphs[0].font.size = Pt(7)
		textbox4 = slide.shapes.add_textbox(Inches(3.66), Inches(1.14), Inches(0.30), Inches(0.22))
		tf4 = textbox4.text_frame
		tf4.paragraphs[0].text = str(appendix_nrows)
		tf4.paragraphs[0].font.size = Pt(7)
		tf4.paragraphs[0].alignment = PP_ALIGN.CENTER
		textbox5 = slide.shapes.add_textbox(Inches(3.66), Inches(1.02), Inches(0.30), Inches(0.22))
		tf5 = textbox5.text_frame
		tf5.paragraphs[0].text = str(preMTB_appendix_nrows)
		tf5.paragraphs[0].font.size = Pt(7)
		tf5.paragraphs[0].alignment = PP_ALIGN.CENTER
		if(TMB_TSO500 != "" and TMB_TSO500 != "NA"):
			if(TMB_TSO500_nr >= 0 and TMB_TSO500_nr <= 5):
				roundshape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(7.07), Cm(3.90), Cm(0.58), Cm(0.58))
				textbox6 = slide.shapes.add_textbox(Inches(2.74), Inches(1.54), Inches(0.32), Inches(0.21))
			if(TMB_TSO500_nr > 5 and TMB_TSO500_nr <= 20):
				roundshape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(8.27), Cm(3.90), Cm(0.58), Cm(0.58))
				textbox6 = slide.shapes.add_textbox(Cm(8.23), Cm(3.90), Cm(0.58), Cm(0.60))
			if(TMB_TSO500_nr > 20):
				roundshape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(10.26), Cm(3.90), Cm(0.58), Cm(0.58))
				textbox6 = slide.shapes.add_textbox(Cm(10.20), Cm(3.95), Cm(0.58), Cm(0.60))
			roundshape.line.color.rgb = RGBColor(255,165,0)
			tf6 = textbox6.text_frame
			tf6.paragraphs[0].text = str(TMB_TSO500_nr)
			tf6.paragraphs[0].font.size = Pt(7)
			tf6.paragraphs[0].font.bold = True
			tf6.paragraphs[0].alignment = PP_ALIGN.CENTER
			tf6.paragraphs[0].font.color.rgb = RGBColor(250,250,250)
		textbox7 = slide.shapes.add_textbox(Inches(6.23), Inches(1.06), Inches(0.97), Inches(0.19))
		tf7 = textbox7.text_frame
		tf7.paragraphs[0].text = splicing
		tf7.paragraphs[0].font.size = Pt(7) 
		tf7.paragraphs[0].alignment = PP_ALIGN.LEFT
		textbox8 = slide.shapes.add_textbox(Inches(6.23), Inches(1.26), Inches(0.97), Inches(0.19))
		tf8 = textbox8.text_frame	
		tf8.paragraphs[0].text = fusion
		tf8.paragraphs[0].font.size = Pt(7)
		tf8.paragraphs[0].alignment = PP_ALIGN.LEFT
		if(index == 1):
			textbox9 = slide.shapes.add_textbox(Inches(5.14), Inches(2.63), Inches(0.53), Inches(0.25))
			tf9 = textbox9.text_frame
			tf9.paragraphs[0].text = TMB_DRUP_str
			tf9.paragraphs[0].font.size = Pt(8)
			tf9.paragraphs[0].alignment = PP_ALIGN.LEFT
			textbox10 = slide.shapes.add_textbox(Inches(3.95), Inches(3.70), Inches(0.88), Inches(0.25))
			tf10 = textbox10.text_frame
			tf10.paragraphs[0].text = TMB_illumina
			tf10.paragraphs[0].font.size = Pt(8)
			tf10.paragraphs[0].alignment = PP_ALIGN.LEFT
			textbox11 = slide.shapes.add_textbox(Inches(4.90), Inches(3.84), Inches(1.14), Inches(0.25))
			tf11 = textbox11.text_frame
			tf11.paragraphs[0].text = MSI_illumina
			tf11.paragraphs[0].font.size = Pt(8)
			tf11.paragraphs[0].alignment = PP_ALIGN.CENTER
		if(index == 5):
			textbox12 = slide.shapes.add_textbox(Inches(6.03), Inches(2.14), Inches(0.59), Inches(0.25))
			tf12 = textbox12.text_frame
			tf12.paragraphs[0].text = TMB_DRUP_str
			tf12.paragraphs[0].font.size = Pt(9)
			tf12.paragraphs[0].alignment = PP_ALIGN.LEFT
			textbox13 = slide.shapes.add_textbox(Inches(6.22), Inches(2.29), Inches(0.88), Inches(0.25))
			tf13 = textbox13.text_frame
			tf13.paragraphs[0].text = TMB_illumina
			tf13.paragraphs[0].font.size = Pt(9)
			tf13.paragraphs[0].alignment = PP_ALIGN.LEFT		
	ppt.save(output_ppt_file)
	return stable_text


def remisse_mail_writer(remisse_file,ipd_no,ipd_consent,DNA_normal_sampleID,RNA_sampleID,extraction_hospital,ipd_material_id,str_TMB_DRUP,TMB_DRUP,stable_text,sample_material,sample_type,sample_list,pipline):
	from docx import Document
	from docx.shared import Pt
	from docx.shared import RGBColor as docRGBColor
	from docx.enum.text import WD_ALIGN_PARAGRAPH
	impress_id = ipd_consent
	sample_type = sample_type.replace("\n", "")
	doc = Document()
	doc.styles['Normal'].font.name = 'Calibri'
	doc.styles['Normal'].font.size = Pt(12)
	pg1 = doc.add_paragraph()
	if(DNA_normal_sampleID != ""):
		if_normal = "med"
	else:
		if_normal = "uten"
	text1 = pg1.add_run("Enhet for studierelatert diagnostikk, Avdeling for patologi, Klinikk for laboratoriemedisin, Oslo Universitetssykehus\nMAL for utsvar av TSO500, tumorprøve " + if_normal + " normal prøve (gDNA)")
	text1.bold = True
	pg1.alignment = WD_ALIGN_PARAGRAPH.CENTER
	pg2 = doc.add_paragraph()
	text2 = pg2.add_run("Merket grønt = fylles inn av molekylærbiolog og/eller patolog manuelt\n\n")
	text2.font.color.rgb = docRGBColor(0,176,80)
	pg2.add_run("PASIENT ID: " + ipd_no + "/" + impress_id + "\n\n")
	pg2.add_run("(lim inn tekst under # i LVMS i diagnose feltet)\n")
	pg2.add_run("##########################################################\n\n")
	if(DNA_normal_sampleID != ""):
		if_normal_sampleID = ") og DNA ekstrahert fra blod:\n\n"
	else:
		if_normal_sampleID = "):\n\n"
	if(RNA_sampleID != ""):
		if_RNA_sampleID = " og RNA "
	else:
		if_RNA_sampleID = " "
	try:
		ipd_material_id_str = ipd_material_id.split(",")
		DNA_material_id = ipd_material_id_str[0].split(":")[1]
		RNA_material_id = ipd_material_id_str[1].split(":")[1]
		if(DNA_material_id == RNA_material_id):
			ipd_material_id_remisse = DNA_material_id
		else:
			ipd_material_id_remisse = ipd_material_id 
	except:
		ipd_material_id_remisse = ipd_material_id
	pg2.add_run("Utvidet genpanelanalyse, TSO500, DNA" + if_RNA_sampleID + "(ekstrahert fra " + extraction_hospital + ") " + ipd_material_id_remisse + ":\n\n")
	if(str_TMB_DRUP == "NA"):
		TMB_string = "Upålitelig, ikke beregnet\n"
	else:
		TMB_string = pronto.get_tmb_string(TMB_DRUP)
	if(stable_text == "Unstable"):
		stable_text = "Ustabil"
	if(stable_text == "Stable"):
		stable_text = "Stabil"
	if(stable_text == "Not reliable"):
		stable_text = "Inkonklusiv"
	pg2.add_run("Tumor mutasjons byrde (TMB) estimat og kategori: " + TMB_string + "Mikrosatellitt (MS) status: " + stable_text + "\nDNA kopitall endringer (estimert kopitall): ") 
	text3 = pg2.add_run("Ingen kopitall av sikker klinisk betydning.\n")
	text3.font.color.rgb = docRGBColor(0,176,80)
	pg2.add_run("Gen-fusjoner: ")
	text4 = pg2.add_run("Ingen funn av sikker klinisk betydning.\n")
	text4.font.color.rgb = docRGBColor(0,176,80)
	pg2.add_run("Somatiske punkt mutasjoner/insersjoner/delesjoner: ")
	text5 = pg2.add_run("Ingen funn av sikker klinisk betydning.\n\n")
	text5.font.color.rgb = docRGBColor(0,176,80)
	pg2.add_run("Se vurdering og vedlegg. \n\n")
	pg2.add_run("(lim inn tekst under # i LVMS i vurderingsfeltet)\n")
	pg2.add_run("##########################################################\n\n")
	year = time.strftime("%Y", time.localtime())
	text6 = pg2.add_run("Funnene er diskutert med overlege XXXX på Mol-MDT-møtet XX.XX." + year + ".\n\nPasienten er diskutert på pre-Mol-MDT-møte XX.XX." + year +".Ettersom det ikke er funn som tilsier utprøvende behandlingsmulighet har man ikke kalt inn behandlende lege til Mol-MDT-møte. Ta kontakt dersom noe er uklart.\n\nKimbane funn som skal følges opp? XXXXXXX. HGVS nomenklatur:\nGEN:ENSTxxx:c.xxx>y:p.AxxxB\n\nDet var dessverre ikke tilstrekkelig mengde og/eller kvalitet av DNA/RNA til at sekvenseringsanalysen kunne gjennomføres.\n\n")
	text6.font.color.rgb = docRGBColor(0,176,80)
	pg2.add_run("Kun funn med klinisk/diagnostisk betydning er rapportert, men se den vedlagte Mol-MDT-rapporten for utfyllende informasjon om testresultatet.\n\nFor teknisk beskrivelse og metodeprinsipp av analyse, vennligst se «TSO 500 genpanelanalyse (utvidet molekylæranalyse)» på nettsiden Metodebok.no under Helse Sør-Øst og OUS.")

	doc.save(remisse_file)


def update_clinical_master_file(InPreD_clinical_data_file,sample_id,if_generate_report,ipd_birth_year,clinical_diagnosis,ipd_gender,ipd_consent,material_id,ipd_collection_year,requisition_hospital,extraction_hospital,tumor_content_nr,batch_nr,pathology_comment,sample_info_comment,encoding_sys):
	global ipd_diagnosis_year
	global runID
	if_exist = False
	new_content = ""
	if(encoding_sys != ""):
		fr = open(InPreD_clinical_data_file, 'r', encoding=encoding_sys)
	else:
		fr = open(InPreD_clinical_data_file, 'r')
	if(str(requisition_hospital) == "0.0"):
		requisition_hospital = "-"
	if(str(extraction_hospital) == "0.0"):
		extraction_hospital = "-"
	for ln in fr:
		if(ln.split('\t')[0] == sample_id):
			if_exist = True
			line = '\t'.join([sample_id, runID, if_generate_report, ipd_birth_year, ipd_diagnosis_year, clinical_diagnosis, ipd_gender, ipd_consent, material_id, ipd_collection_year, requisition_hospital, extraction_hospital, str(tumor_content_nr), batch_nr, pathology_comment, sample_info_comment + '\n'])
			new_line = ln.replace(ln,line)
			new_content = new_content + new_line
		else:
			new_content = new_content + ln
	fr.close()
	if(encoding_sys != ""):
		fa = open(InPreD_clinical_data_file, 'a', encoding=encoding_sys)
	else:
		fa = open(InPreD_clinical_data_file, 'a')
	if(if_exist == False):
		line = '\t'.join([sample_id, runID, if_generate_report, ipd_birth_year, ipd_diagnosis_year, clinical_diagnosis, ipd_gender, ipd_consent, material_id, ipd_collection_year, requisition_hospital, extraction_hospital, str(tumor_content_nr), batch_nr, pathology_comment, sample_info_comment + '\n'])
		if(encoding_sys != ""):
			fa = open(InPreD_clinical_data_file, 'a', encoding=encoding_sys)
		else:
			fa = open(InPreD_clinical_data_file, 'a')
		fa.write(line)
		fa.close()
	else:
		if(encoding_sys != ""):
			fw = open(InPreD_clinical_data_file, 'w', encoding=encoding_sys)
		else:
			fw = open(InPreD_clinical_data_file, 'w')
		fw.write(new_content)
		fw.close()


def update_clinical_tsoppi_file(InPreD_clinical_tsoppi_data_file,sample_id,if_generate_report,ipd_birth_year,clinical_diagnosis,ipd_gender,ipd_consent,material_id,ipd_collection_year,requisition_hospital,extraction_hospital,tumor_content_nr,batch_nr,sample_material,sample_type,tumor_type,TMB_DRUP,TMB_TSO500,MSI_TSO500,pipline,pathology_comment,sample_info_comment):
	if_exist = False
	assay_name = ""
	nucleicacid = ""
	RNA_DNA_tumor_normal = ""
	global ipd_diagnosis_year
	global runID
	try:
		sample_type = sample_type.replace("\n", "")
	except:
		sample_type = ""
	if(str(requisition_hospital) == "0.0"):
		requisition_hospital = "-"
	if(str(extraction_hospital) == "0.0"):
		extraction_hospital = "-"
	if(pipline != "-" and pipline != ""):
		pipline = pipline.split(": ")[1]
	new_content = ""
	with open(InPreD_clinical_tsoppi_data_file, 'r') as fr:
		for ln in fr:
			if(ln.split('\t')[0] == sample_id):
				if_exist = True
				line = '\t'.join([sample_id, runID, if_generate_report, ipd_birth_year, ipd_diagnosis_year, clinical_diagnosis, ipd_gender, ipd_consent, material_id, ipd_collection_year, requisition_hospital, extraction_hospital, str(tumor_content_nr), batch_nr, sample_material, sample_type, tumor_type, str(TMB_DRUP), TMB_TSO500, MSI_TSO500, pipline, pathology_comment, sample_info_comment + '\n'])
				new_line = ln.replace(ln,line)
				new_content = new_content + new_line
			else:
				new_content = new_content + ln
	fr.close()
	if(if_exist == False):
		line = '\t'.join([sample_id, runID, if_generate_report, ipd_birth_year, ipd_diagnosis_year, clinical_diagnosis, ipd_gender, ipd_consent, material_id, ipd_collection_year, requisition_hospital, extraction_hospital, str(tumor_content_nr), batch_nr, sample_material, sample_type, tumor_type, str(TMB_DRUP), TMB_TSO500, MSI_TSO500, pipline, pathology_comment, sample_info_comment + '\n'])
		with open(InPreD_clinical_tsoppi_data_file, 'a') as fa:
			fa.write(line)
		fa.close()
	else:
		with open(InPreD_clinical_tsoppi_data_file, 'w') as fw:
			fw.write(new_content)
		fw.close()

def pdf_page_image_to_ppt(pdf_file,output_ppt_file,pages,width_scale,height_scale):
	ppt = Presentation(output_ppt_file)
	slide = ppt.slides.add_slide(ppt.slide_layouts[6])
	image_width = ppt.slide_width * width_scale
	image_height = ppt.slide_height * height_scale
	for i in range(0, len(pages)):
		images = convert_from_path(pdf_file, first_page=pages[i], last_page=pages[i])
		if not images:
			logging.warning("The page number {} in {} is not found!".format(pages[i], pdf_file))
			continue
		image_path = "temp_page_{}.jpg".format(pages[i])
		images[0].save(image_path, 'JPEG')
		if(len(pages) == 1):
			image_top = ppt.slide_height/4
		else:
			image_top = image_height*i
		slide.shapes.add_picture(image_path, left=0, top=image_top, width=image_width, height=image_height)
		os.remove(image_path)
	ppt.save(output_ppt_file)

def usage(exit_status = 0):
	print ("""Usage: python3  %s
        This script is a tool used to generate the paitent report based on the TSO500 analysis results and the personal intomation from the clinical data in In/InPreD_PRONTO_metadata.txt,
	and update the SOPPI results into the file Out/InPreD_PRONTO_metadata_tsoppi.txt when the reports are generated.
	This script could also fill the patient personal information into the clinical data file with the MTF files under the foder In/MTF/. (This fuction currently is only used by OUS)
	To run this script tool in your computer with python3, it will read the clinical data from In/InPreD_PRONTO_metadata.txt and generate reports for the Sample_id with Create_report==Y:
	
	python3 InPreD_PRONTO.py

	Extra parameters for OUS:
	-c, --clinical_file Fill the patient personal information into the clinical data file: InPreD_PRONTO_metadata.txt with the MTF files under the foder In/MTF/
	python3 InPreD_PRONTO.py -D <DNA_sampleID> -r <runID> -c
	or:
	python3 InPreD_PRONTO.py --DNAsampleID=<DNA_sampleID> --runID=<runID> --clinicalFile
	-m, --mail_draft Generate the Remisse_draft.docx file with report:
	python3 InPreD_PRONTO.py -m
	or:
	python3 InPreD_PRONTO.py --mailDraft

	This script will create sub-folder with runID/IPDXXX in Out/, move the IPD_Material file into it and generate all the results files under that sub-folder.
	 
	-h, --help See this help information and exit.
        """ % sys.argv[0])
	sys.exit(exit_status)


def main(argv):
	global runID
	global DNA_sampleID
	global RNA_sampleID
	global ipd_birth_year
	global ipd_clinical_diagnosis
	global ipd_gender
	global ipd_consent
	global ipd_material_id
	global DNA_material_id
	global RNA_material_id
	global ipd_collection_year
	global pathology_comment
	global requisition_hospital
	global extraction_hospital
	global sample_material
	global sample_type
	global tumor_type
	global batch_nr
	global tumor_content_nr
	global sample_info_comment
	global TMB_DRUP
	global str_TMB_DRUP
	global TMB_TSO500
	global MSI_TSO500
	global pipline
	runID_RNA = ""
	runID_DNA = ""
	DNA_normal_sampleID = ""
	remisse_mail = False
	update_clinical_file = False
	target_cod_region = 0
	tumor_content = "XX"
	ipd_material_file_2023 = ""
	ipd_material_file_new = ""
	try:
		opts, args = getopt.getopt(sys.argv[1:], "hr:D:mc", ["help", "runID=", "DNAsampleID=", "mailDraft", "clinicalFile"])
	except getopt.GetoptError:
		usage(1)

	for opt, arg in opts:
		if opt in ("-h", "--help"):
			usage()
		elif opt in ("-r", "--runID"):
			runID = arg
		elif opt in ("-D", "--DNAsampleID"):
			DNA_sampleID = arg
			ipd_no = DNA_sampleID.split('-')[0]
		elif opt in ("-m", "--mailDraft"):
			remisse_mail = True
		elif opt in ("-c", "--clinicalFile"):
			update_clinical_file = True
	runID_DNA = runID
	DNA_sampleID_format = '^IP[A-Z]\\d{4}-D(\\d|X){2}-[A-z](\\d|X){2}-[A-z](\\d|X){2}'

	base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
	config_file = base_dir + "/Config/configure_PRONTO.ini"
	InPreD_clinical_data_file = base_dir + "/In/InPreD_PRONTO_metadata.txt"
	output_path_root = base_dir + "/Out/"
	InPreD_clinical_tsoppi_data_file = base_dir + "/Out/InPreD_PRONTO_metadata_tsoppi.txt"
	cfg = ConfigParser()
	cfg.read(config_file)
	inpred_node = cfg.get("INPUT", "inpred_node")
	data_path = cfg.get("INPUT", "data_path")
	encoding_sys = cfg.get("INPUT", "encoding_sys")
	filter_col_nu_config = int(cfg.get("INPUT", "filter_col_nu"))

	if not os.path.exists(InPreD_clinical_data_file):
		print ("""      Error:
		The InPreD clinical file InPreD_PRONTO_metadata.txt does not exist!
		""")
		sys.exit(0)
	if(update_clinical_file == True):
		material_file_version = cfg.get("INPUT", "material_file_version")
		if(int(material_file_version) <= 2023):
			ipd_material_file_2023 = base_dir + "/In/MTF/" + ipd_no[:3] + '-' + ipd_no[3:] + "_Material Transit Form InPreD NGS.xlsx"
		if(int(material_file_version) > 2023):
			ipd_material_file_new = base_dir + "/In/MTF/" + ipd_no[:3] + '-' + ipd_no[3:] + "_Material Transit Form InPreD NGS_" + material_file_version + ".xlsx"
		if not(os.path.exists(ipd_material_file_new) or os.path.exists(ipd_material_file_2023)):
			print ("""Error: IPD Material Transit Form InPreD NGS file does not exit under the MTF dir. PRONTO meta file could not be updated with patient personal information by parameter -c of this script!""")
			sys.exit(0)
		if not(re.fullmatch(DNA_sampleID_format, DNA_sampleID)):
			print("Warning: " + DNA_sampleID + " does not fit for the sample id format!")
		else:
			sample_list_file = pronto.glob_tsoppi_file(True, data_path, runID_DNA, DNA_sampleID, 'sample_list.tsv')
			for line in open(sample_list_file):
				if(line.startswith("RNA_tumor")):
					RNA_sampleID = line.split('\t')[1]
			if os.path.exists(ipd_material_file_new):
				get_patient_info_from_MTF_new(ipd_material_file_new,ipd_no,DNA_sampleID,RNA_sampleID)
			if os.path.exists(ipd_material_file_2023):
				get_patient_info_from_MTF_2023(ipd_material_file_2023,ipd_no,DNA_sampleID,RNA_sampleID)
			if_generate_report = "Y"
			update_clinical_master_file(InPreD_clinical_data_file,DNA_sampleID,if_generate_report,ipd_birth_year,ipd_clinical_diagnosis,ipd_gender,ipd_consent,DNA_material_id,ipd_collection_year,requisition_hospital,extraction_hospital,tumor_content_nr,batch_nr,pathology_comment,sample_info_comment,encoding_sys)
			print("Clinical data is added into PRONTO meta file for sample: " + DNA_sampleID)
			if(RNA_sampleID != ""):
				if_generate_report = "-"
				update_clinical_master_file(InPreD_clinical_data_file,RNA_sampleID,if_generate_report,ipd_birth_year,ipd_clinical_diagnosis,ipd_gender,ipd_consent,RNA_material_id,ipd_collection_year,requisition_hospital,extraction_hospital,tumor_content_nr,batch_nr,pathology_comment,sample_info_comment,encoding_sys)
				print("Clinical data is added into PRONTO meta file for sample: " + RNA_sampleID)
			sys.exit(0)
	ppt_nr = 0
	if(encoding_sys != ""):
		meta_file = open(InPreD_clinical_data_file, encoding=encoding_sys)
	else:
		meta_file = open(InPreD_clinical_data_file)
	for ln in meta_file:
		if not(ln.startswith("#") or ln == ""):
			if(ln.split('\t')[2] == "Y"):
				ln = ln.replace("\n", "")
				global extra_path
				RNA_sampleID = ""
				sample_list = []
				ln = ln + '\t' * (15-ln.count('\t'))
				[DNA_sampleID, runID_DNA, _, ipd_birth_year, ipd_diagnosis_year, ipd_clinical_diagnosis_meta, ipd_gender, ipd_consent, DNA_material_id, ipd_collection_year, requisition_hospital, extraction_hospital, tumor_content_nr, batch_nr, pathology_comment, sample_info_comment] = ln.split('\t')
				runID = runID_DNA
				if not(re.fullmatch(DNA_sampleID_format, DNA_sampleID)):
					print("Warning: " + DNA_sampleID + " does not fit for the sample id format!")
				try:
					ipd_age = str(int(time.strftime("%Y", time.localtime())) - int(ipd_birth_year))
				except:
					ipd_age = "-"
				ipd_no = DNA_sampleID.split('-')[0]
				try:
					ipd_clinical_diagnosis_ppt = ipd_clinical_diagnosis_meta.split("(")[0] + "\n(" + ipd_clinical_diagnosis_meta.split("(")[1]
				except:
					ipd_clinical_diagnosis_ppt = ipd_clinical_diagnosis_meta
				output_path = output_path_root + runID + "/" + DNA_sampleID + "/"
				extra_path = output_path + "extra_files"
				output_file_preMTB_table_path = output_path + DNA_sampleID
				sample_list_file = pronto.glob_tsoppi_file(True, data_path, runID_DNA, DNA_sampleID, 'sample_list.tsv')
				data_file_small_variant_table = pronto.glob_tsoppi_file(False, data_path, runID_DNA, DNA_sampleID, '{}_small_variant_table_forQC.tsv'.format(DNA_sampleID))
				CNV_overview_plots_pdf = pronto.glob_tsoppi_file(True, data_path, runID_DNA, DNA_sampleID, '{}_CNV_overview_plots.pdf'.format(DNA_sampleID))
				if not data_file_small_variant_table:
					data_file_small_variant_table = pronto.glob_tsoppi_file(True, data_path, runID_DNA, DNA_sampleID, '{}_small_variant_table.tsv'.format(DNA_sampleID))
				data_file_cnv_overview_plots = pronto.glob_tsoppi_file(True, data_path, runID_DNA, DNA_sampleID, '{}_CNV_overview_plots.pdf'.format(DNA_sampleID))

				if not os.path.exists(data_file_small_variant_table):
					print ("Error:  The data input file " + data_file_small_variant_table + " does not exist!")
					sys.exit(0)
				if not os.path.exists(output_path):
					os.makedirs(output_path)
				if not os.path.exists(extra_path):
					os.makedirs(extra_path)
				for line in open(sample_list_file):
					if not(line.startswith("#")):
						Sample_Type = line.split('\t')[0]
						Sample_ID = line.split('\t')[1]
						sample_list_str = Sample_Type + ": " + Sample_ID
						sample_list.append(sample_list_str)
					if(line.startswith("DNA_tumor")):
						DNA_run_dir = line.split('\t')[4]
					if(line.startswith("DNA_normal")):
						DNA_normal_sampleID = line.split('\t')[1]
					if(line.startswith("RNA_tumor")):
						RNA_run_dir = line.split('\t')[4]
						RNA_sampleID = line.split('\t')[1]		
					else:
						runID_RNA = ""
						RNA_sampleID = ""
						RNA_material_id = ""
				if(RNA_sampleID != ""):
					if (DNA_run_dir == RNA_run_dir):
						runID_RNA = runID_DNA
					else:
						# split according to suffix (_TSO_500)_LocalApp_results and use on RNA_run_dir basename (last element in path)
						tokens = re.split('(_TSO_500)?_LocalApp_results', os.path.basename(RNA_run_dir))
						runID_RNA = tokens[0]
					RNA_material_id = get_RNA_material_id(InPreD_clinical_data_file,RNA_sampleID,encoding_sys)
					ipd_material_id = "DNA:" + DNA_material_id + ",RNA:" + RNA_material_id
				else:
					ipd_material_id = "DNA:" + DNA_material_id

				for line in open(data_file_small_variant_table):
					if(line.startswith("#") and "Version string:" in line):
						pipline = line.split('\t')[0]
						pipline = pipline.split('] ')[1]
						pipline = pipline.split('\n')[0]
					if(line.startswith('#') and "Specified tumor purity" in line):
						if("not specified" in line or "a default value of 0.5 is being used" in line):
							tumor_content = "XX"
						else:
							tumor_content_float = Decimal(line.split(':')[1]).quantize(Decimal('0.00'))
							tumor_content = '~' + str(int(tumor_content_float*100)) + '%'
					if(line.startswith('#') and "Size of the target coding region" in line):
						target_cod_region = float(line.split(':')[1])
				for i in range(0,filter_col_nu_config+1):
					filter_section = str(i)
					if(filter_section == "0"):
						concat_data = pandas.DataFrame()
						topfilters = pronto.parse_topfilter(cfg, output_file_preMTB_table_path)
						for filter in topfilters:
							for filter_column in filter['filter_columns']:
								small_variant_data = pandas.read_csv(data_file_small_variant_table, sep='\t', comment='#')
								try:
									small_variant_data = pronto.filter_small_variant_data(small_variant_data, DNA_sampleID, filter_column, filter['key_word'])
								except ValueError:
									sys.exit(1)
							if filter['filter_column_add']:
								small_variant_data[filter['filter_column_add']] = 'Yes'
							if filter['min_depth_tumor_dna']:
								small_variant_data = small_variant_data[small_variant_data['Depth_tumor_DNA'] >= int(filter['min_depth_tumor_dna'])]
							small_variant_data.to_csv(filter['table_output_path'], sep='\t', index=False)
							concat_data = pandas.concat([concat_data, small_variant_data])
						
						concat_data = concat_data.drop_duplicates()
						for i, irow in concat_data.iterrows():
							for j, jrow in concat_data.iterrows():
								# avoid comparing the same rows twice
								if i >= j:
									continue

								# get the difference between the two rows and check if the only different is the filter column with NA value
								diff = irow.compare(jrow)
								if len(diff) == 1 and diff.isna().any().any():
									# if yes, set the filter column with NA value to empty string in both rows
									concat_data.at[i, diff.index.tolist()[0]] = ''
									concat_data.at[j, diff.index.tolist()[0]] = ''
						# remove duplicates
						concat_data = concat_data.drop_duplicates()
						top_filter_output_file = "{}_preMTB_workingTable.txt".format(output_file_preMTB_table_path)
						concat_data.write_csv(top_filter_output_file, sep='\t', index=False)
						continue

					if(filter_section == "1"):
						MTB_format = True
					filter_column = cfg.get("FILTER"+filter_section, "filter_column")
					key_word = cfg.get("FILTER"+filter_section, "key_word")
					columns = cfg.get("FILTER"+filter_section, "columns")
					output_table = cfg.get("FILTER"+filter_section, "output_table")
					output_table_file_config_pre = output_file_preMTB_table_path + "_" + output_table + "_pre.txt"
					output_table_file_config = output_file_preMTB_table_path + "_" + output_table + ".txt"
					if(DNA_normal_sampleID != ""):
						columns = columns + ",AF_normal_DNA"
					all_data = read_tsv_col(top_filter_output_file,filter_column,key_word,columns,MTB_format)
					write_exl(output_table_file_config_pre,all_data)
					clear_blank_line(output_table_file_config_pre,output_table_file_config)
					MTB_format = False

				ppt_template = base_dir + "/In/Template/InPreD_MTB_template.pptx"
				DNA_variant_summary_file = pronto.glob_tsoppi_file(True, data_path, runID_DNA, '{}_variant_summary.tsv'.format(runID_DNA))
				if(runID_RNA != ""):
					RNA_variant_summary_file = pronto.glob_tsoppi_file(True, data_path, runID_RNA, '{}_variant_summary.tsv'.format(runID_RNA))
				else:
					RNA_variant_summary_file = ""
				output_ppt_file = output_path + DNA_sampleID + "_MTB_report.pptx"
				DNA_image_path = pronto.glob_tsoppi_file(True, data_path, runID_DNA, DNA_sampleID)
				if(RNA_sampleID != ""):
					RNA_image_path = DNA_image_path
				else:
					RNA_image_path = ""

				file = os.path.split(data_file_small_variant_table)[1]
				try:
					sample_type_string = file.split('-')[2]
					sample_type_short = sample_type_string[0:1]
					sample_type_list = {'M': 'Metastasis', 'T': 'Primary Tumor', 'C': 'Cell-line', 'N': 'Normal/Control', 'P': 'Primary tumor\n naive', 'p': 'Primary tumor\n post-treatment', 'R': 'Regional met\n naive', 'r': 'Regional met\n post-treatment', 'D': 'Distal met\n naive', 'd': 'Distal met\n post-treatment', 'L': 'Liquid', 'E': 'naive', 'e': 'post treatment', 'A': 'post allo transplant', 'X': 'Unknown'}
					sample_type = sample_type_list.get(sample_type_short)
				except:
					sample_type = ""
				try:
					sample_material_string = file.split('-')[3]
					sample_material_short = sample_material_string[0:1]
					sample_material_list = {'F': 'Fresh Frozen', 'A': 'Archived FFPE', 'B': 'Blood', 'C': 'Cytology', 'M': 'Fresh bone marrow',  'E': 'Extramedullary','S': 'Buccal swab (normal)', 'X': 'Unspecified'}
					sample_material = sample_material_list.get(sample_material_short)
				except:
					sample_material = ""
				try:
					tumor_type_no = sample_material_string[1:3]
					tumor_type_list = {'00': 'Cancer origo incerta', '01': 'Adrenal Gland', '02': 'Ampulla of Vater', '03': 'Biliary Tract', '04': 'Bladder/Urinary Tract', '05': 'Bone', '06': 'Breast', '07': 'Cervix', '08': 'CNS/Brain', '09': 'Colon/Rectum', '10': 'Esophagus/Stomach', '11': 'Eye', '12': 'Head and Neck', '13': 'Kidney', '14': 'Liver', '15': 'Lung', '16': 'Lymphoid', '17': 'Myeloid', '18': 'Ovary/Fallopian Tube', '19': 'Pancreas', '20': 'Peripheral Nervous System', '21': 'Peritoneum', '22': 'Pleura', '23': 'Prostate', '24': 'Skin', '25': 'Soft Tissue', '26': 'Testis', '27': 'Thymus', '28': 'Thyroid', '29': 'Uterus', '30': 'Vulva/Vagina', 'XX': 'Not available'}
					tumor_type = tumor_type_list.get(tumor_type_no)
				except:
					tumor_type = ""
                
				min_AF_tumor_DNA = float(cfg.get("TMB", "min_AF_tumor_DNA"))
				min_depth_tumor_DNA = int(cfg.get("TMB", "min_depth_tumor_DNA"))
				TMB_filter_column = cfg.get("TMB", "TMB_filter_column")
				TMB_filter_key_word = cfg.get("TMB", "TMB_filter_key_word")
				TMB_coding_file_pre = output_path + DNA_sampleID + "_TMB_coding_pre.txt"
				TMB_coding_file = output_path + DNA_sampleID + "_TMB_coding.txt"
				TMB_DURP_coding_file_pre = output_path + DNA_sampleID + "_TMB_DURP_coding_pre.txt"
				TMB_DURP_coding_file = output_path + DNA_sampleID + "_TMB_DURP_coding.txt"
				TMB_DRUP_filter_key_word = cfg.get("TMB", "TMB_DRUP_filter_key_word")

				TMB_coding_data = read_tsv(top_filter_output_file,TMB_filter_column,TMB_filter_key_word)
				write_exl(TMB_coding_file_pre,TMB_coding_data)
				clear_blank_line(TMB_coding_file_pre,TMB_coding_file)
				TMB_DRUP_coding_data = read_tsv(top_filter_output_file,TMB_filter_column,TMB_DRUP_filter_key_word)
				write_exl(TMB_DURP_coding_file_pre,TMB_DRUP_coding_data)
				clear_blank_line(TMB_DURP_coding_file_pre,TMB_DURP_coding_file)

				rows_preMTB_AFTumor = 0
				for line in open(TMB_coding_file):
					line_cells = line.split('\t')
					if(line.startswith("Sample_ID")):
						for col in range(len(line_cells)):
							if(line_cells[col] == "AF_tumor_DNA"):
								col_AF_tumor = col
							if(line_cells[col] == "Depth_tumor_DNA"):
								col_Depth_tumor = col
					else:
						AF_tumor_DNA = float(line_cells[col_AF_tumor])
						Depth_tumor_DNA = int(line_cells[col_Depth_tumor])
						if(AF_tumor_DNA >= min_AF_tumor_DNA and Depth_tumor_DNA >= min_depth_tumor_DNA):
							rows_preMTB_AFTumor += 1
		
				rows_TMB_DRUP_AFTumor = 0
				for line in open(TMB_DURP_coding_file):
					line_cells = line.split('\t')
					if(line.startswith("Sample_ID")):
						for col in range(len(line_cells)):
							if(line_cells[col] == "AF_tumor_DNA"):
								col_DRUP_AF_tumor = col
							if(line_cells[col] == "Depth_tumor_DNA"):
								col_DRUP_Depth_tumor = col
					else:
						DRUP_AF_tumor_DNA = float(line_cells[col_DRUP_AF_tumor])
						DRUP_Depth_tumor_DNA = int(line_cells[col_DRUP_Depth_tumor])
						if(DRUP_AF_tumor_DNA >= min_AF_tumor_DNA and DRUP_Depth_tumor_DNA >= min_depth_tumor_DNA):
							rows_TMB_DRUP_AFTumor += 1

				if(target_cod_region == 0):
					TMB_DRUP = -1
					TMB_DRUP_str = "-1"
				else:
					TMB_DRUP = round(rows_TMB_DRUP_AFTumor/target_cod_region)
					TMB_DRUP_str = str(rows_TMB_DRUP_AFTumor) + '/' + str(target_cod_region)

				update_ppt_template_data(inpred_node,ipd_no,ipd_gender,ipd_age,ipd_diagnosis_year,DNA_material_id,RNA_material_id,ipd_consent,requisition_hospital,pathology_comment,ipd_clinical_diagnosis_ppt,tumor_type,sample_type,sample_material,sample_info_comment,pipline,tumor_content,ppt_template,output_ppt_file)

				insert_image_to_ppt(DNA_sampleID,DNA_normal_sampleID,RNA_sampleID,DNA_image_path,RNA_image_path,output_ppt_file)

                		# Insert tables into PP file:
				slide8_table_header = ["Gene_symbol", "Genomic_coordinates_in_hg19_build", "Ensembl_transcript_ID", "Exon_number", "Protein_change_short", "HGVS_syntax", "Change_summary", "Coding_status", "Read_depth(variant reads/total reads)", "AF_tumor_DNA", "Filter_rescued"]
				if(DNA_normal_sampleID != ""):
					slide6_table_header = ["Gene_symbol", "Protein_change_short", "Coding_status", "AF_tumor_DNA", "AF_normal_DNA"]
				else:
					slide6_table_header = ["Gene_symbol", "Protein_change_short", "Coding_status", "AF_tumor_DNA"]
        
                		# Slide2, slide6 and slide7 right side table: Variants that alter protein coding sequence
				slide6_table_data_file = output_file_preMTB_table_path + "_AllReporVariants_AltProtein.txt" 
				slide6_table_ppSlide = [2,6,7]
				slide6_table_name = "Variants that alter protein coding sequence "
				slide6_header_left = 7.36
				slide6_header_top = 0.82
				slide6_header_width = 2.55
				slide6_table_left = 7.23
				slide6_table_top = 1.06
				slide6_table_width = 2.76
				slide6_table_height = 1.63
				slide6_table_font_size = 7
				if_print_rowNo = False
				for table_index in slide6_table_ppSlide:
					slide6_table_nrows = insert_table_to_ppt(slide6_table_data_file,table_index,slide6_table_name,slide6_header_left,slide6_header_top,slide6_header_width,slide6_table_left,slide6_table_top,slide6_table_width,slide6_table_height,slide6_table_font_size,slide6_table_header,output_ppt_file,if_print_rowNo,[],table_max_rows_per_slide=None)
				output_file_preMTB_AppendixTable = output_file_preMTB_table_path + "_preMTBTable_Appendix.txt"
				output_table_file_filterResults_AllReporVariants_CodingRegion = output_file_preMTB_table_path + "_AllReporVariants_CodingRegion.txt"
				stable_text = update_ppt_variant_summary_table(slide6_table_nrows,DNA_sampleID,RNA_sampleID,TMB_DRUP,TMB_DRUP_str,DNA_variant_summary_file,RNA_variant_summary_file,output_file_preMTB_AppendixTable,output_table_file_filterResults_AllReporVariants_CodingRegion,output_ppt_file)
				output_file_sequence_summary_table = output_file_preMTB_table_path + "_sequence_summary.txt"
                
				# Slide8 Table: Sequence data summary: Variants alter protein code*
				slide8_table_data_file = output_file_sequence_summary_table
				slide8_table_ppSlide = 9
				slide8_table_name = "Sequence data summary for {}: Variants that alter protein code and splice sites.  TSOPPI Version string: {} ".format(DNA_sampleID, pipline)
				slide8_header_left = 0.25
				slide8_header_top = 0.27
				slide8_header_width = 8.98
				slide8_table_left = 0.15
				slide8_table_top = 0.55
				slide8_table_width = 9.19
				slide8_table_height = 2.81
				slide8_table_font_size = 7
				if_print_rowNo = True
				table8_column_width = [0.54, 0.96, 0.96, 0.51, 0.73, 1.12, 2.26, 0.79, 0.81, 0.53, 0.53]
				table_max_rows_per_slide = int(cfg.get("INPUT", "table_max_rows_per_slide"))
				_ = insert_table_to_ppt(slide8_table_data_file,slide8_table_ppSlide,slide8_table_name,slide8_header_left,slide8_header_top,slide8_header_width,slide8_table_left,slide8_table_top,slide8_table_width,slide8_table_height,slide8_table_font_size,slide8_table_header,output_ppt_file,if_print_rowNo,table8_column_width,table_max_rows_per_slide)

				# Insert the CNV_overveiw_plots pictures A2, B3 and C1 into report.
				A2_to_extract=[2]
				pdf_page_image_to_ppt(CNV_overview_plots_pdf,output_ppt_file,A2_to_extract,width_scale=1,height_scale=0.5)
				if(DNA_normal_sampleID != ""):
					B3_C1_to_extract = [5, 6]
				else:
					B3_C1_to_extract = [4, 5]
				pdf_page_image_to_ppt(CNV_overview_plots_pdf,output_ppt_file,B3_C1_to_extract,width_scale=1,height_scale=0.5)

        			# Change slides order.
				ppt = Presentation(output_ppt_file)
				slides = ppt.slides._sldIdLst
				slides_list = list(slides)
				slides.remove(slides_list[7])
				slides.append(slides_list[7])
				ppt.save(output_ppt_file)
				print("Generate report for " + DNA_sampleID)
				ppt_nr += 1

        			# Move TXT files generated by this script into extra_files folder.
				txt_files = os.listdir(output_path)
				for txt_file in txt_files:
					if(txt_file.endswith('.txt')):
						txt_file_path = os.path.join(output_path,txt_file)
						txt_file_extra_path = os.path.join(extra_path,txt_file)
						if os.path.exists(txt_file_extra_path):
							os.remove(txt_file_extra_path)
						shutil.move(txt_file_path, extra_path)

				# Move small variant data file and plots into the report folder.
				data_file_small_variant_table_cp = output_path + DNA_sampleID + "_small_variant_table_forQC.tsv"
				shutil.copyfile(data_file_small_variant_table, data_file_small_variant_table_cp)
				data_file_cnv_overview_plots_cp = output_path + DNA_sampleID + "_CNV_overview_plots.pdf"
				shutil.copyfile(data_file_cnv_overview_plots, data_file_cnv_overview_plots_cp)

				if(remisse_mail == True):
					remisse_file = output_path + ipd_no + "_Remisse_draft.docx"
					remisse_mail_writer(remisse_file,ipd_no,ipd_consent,DNA_normal_sampleID,RNA_sampleID,extraction_hospital,ipd_material_id,str_TMB_DRUP,TMB_DRUP,stable_text,str(sample_material),sample_type,sample_list,pipline)
				# Move MTF file into extra_files folder if it exists.
				if os.path.exists(ipd_material_file_new):
					move_ipd_material_file = shutil.move(ipd_material_file_new, extra_path)
				if os.path.exists(ipd_material_file_2023):
					move_ipd_material_file = shutil.move(ipd_material_file_2023, extra_path)
					update_clinical_tsoppi_file(InPreD_clinical_tsoppi_data_file,DNA_sampleID,DNA_if_generate_report,ipd_birth_year,ipd_clinical_diagnosis,ipd_gender,ipd_consent,DNA_material_id,ipd_collection_year,requisition_hospital,extraction_hospital,tumor_content_nr,batch_nr,str(sample_material),sample_type,str(tumor_type),str_TMB_DRUP,TMB_TSO500,MSI_TSO500,pipline,pathology_comment,sample_info_comment)
					if(RNA_sampleID != ""):
						RNA_if_generate_report = "-"
						RNA_str_TMB_DRUP = "-"
						RNA_TMB_TSO500 = "-"
						RNA_MSI_TSO500 = "-"
						RNA_pipline = "-"
						update_clinical_tsoppi_file(InPreD_clinical_tsoppi_data_file,RNA_sampleID,RNA_if_generate_report,ipd_birth_year,ipd_clinical_diagnosis,ipd_gender,ipd_consent,RNA_material_id,ipd_collection_year,requisition_hospital,extraction_hospital,tumor_content_nr,batch_nr,sample_material,sample_type,tumor_type,RNA_str_TMB_DRUP,RNA_TMB_TSO500,RNA_MSI_TSO500,RNA_pipline,pathology_comment,sample_info_comment)
	if(ppt_nr > 1):	
		print("Go through the InPreD_PRONTO_metadata file, " + str(ppt_nr) +" reports are generated.")
	else:
		print("Go through the InPreD_PRONTO_metadata file, " + str(ppt_nr) +" report is generated.")

if __name__ == '__main__':
    main(sys.argv[1:])
