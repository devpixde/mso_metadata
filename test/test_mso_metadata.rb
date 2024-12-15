# frozen_string_literal: true

require "test_helper"

class TestMsoMetadata < Minitest::Test
  def test_that_it_has_a_version_number
    refute_nil ::MsoMetadata::VERSION
  end

  def test_ppt
    metadata = MsoMetadata.read "#{__dir__}/sample_files/ppt_meta.pptx"
    assert_equal metadata[:core][:title], 'Eine PPT mit Metadaten'
    assert_equal metadata[:core][:subject], 'Test ppt'
    assert_equal metadata[:core][:creator], 'Ingo Klemm'
    assert_equal metadata[:core][:description], 'Kommentar zur PPT'
    assert_equal metadata[:core][:lastModifiedBy], 'Ingo Klemm'
    assert_equal metadata[:core][:revision], '1'
    assert_equal metadata[:core][:category], "'Test"  # misspelling but ok for testing
    assert_equal metadata[:core][:created], '2024-12-15T10:37:05Z'
    assert_equal metadata[:core][:modified], '2024-12-15T10:38:24Z'
    assert_equal metadata[:core][:application], 'Microsoft Office PowerPoint'
    assert_equal metadata[:core][:company], ''

    assert_equal metadata[:custom]["Test"], "Dies ist ein Test"
    assert_equal metadata[:custom]["Dokumentnummer"], "4711"
  end

  def test_word1
    metadata = MsoMetadata.read "#{__dir__}/sample_files/word_meta1.docx"
    assert_equal metadata[:core][:title], 'Word doc title'
    assert_equal metadata[:core][:subject], 'Topic'
    assert_equal metadata[:core][:creator], 'Ingo Klemm'
    assert_equal metadata[:core][:description], 'Hier kommt ein Kommentar'
    assert_equal metadata[:core][:lastModifiedBy], 'Ingo Klemm'
    assert_equal metadata[:core][:revision], '3'
    assert_equal metadata[:core][:category], 'Kategorie'
    assert_equal metadata[:core][:created], '2024-12-15T10:29:00Z'
    assert_equal metadata[:core][:modified], '2024-12-15T10:33:00Z'
    assert_equal metadata[:core][:application], 'Microsoft Office Word'
    assert_equal metadata[:core][:company], 'Private'

    assert_equal metadata[:custom]['Veröffentlicht'], false
    assert_equal metadata[:custom]['Eine Zahl'], 200
    assert_equal metadata[:custom]['Ein Datum'], "1966-12-18T10:00:00Z"
    assert_equal metadata[:custom]['Status'], "WIP"
    assert_equal metadata[:custom]['Ablage'], "Test"
  end

  def test_word2
    metadata = MsoMetadata.read "#{__dir__}/sample_files/word_meta2.docx"
    assert_equal metadata[:core][:title], ''
    assert_equal metadata[:core][:subject], ''
    assert_equal metadata[:core][:creator], 'Ingo Klemm'
    assert_equal metadata[:core][:description], ''
    assert_equal metadata[:core][:lastModifiedBy], 'Ingo Klemm'
    assert_equal metadata[:core][:revision], '3'
    assert_nil metadata[:core][:category]
    assert_equal metadata[:core][:created], '2024-12-15T10:33:00Z'
    assert_equal metadata[:core][:modified], '2024-12-15T10:34:00Z'
    assert_equal metadata[:core][:application], 'Microsoft Office Word'
    assert_equal metadata[:core][:company], ''

    assert_equal metadata[:custom]['Sprache'], 'bayerisch'
    assert_equal metadata[:custom]['Lorem'], true
  end


  def test_xlsx
    metadata = MsoMetadata.read "#{__dir__}/sample_files/excel_meta.xlsx"
    assert_equal metadata[:core][:title], 'Eine Excel Datei'
    assert_equal metadata[:core][:subject], 'zum Thema'
    assert_equal metadata[:core][:creator], 'Ingo Klemm'
    assert_equal metadata[:core][:description], 'Dies ist ein Test'
    assert_equal metadata[:core][:lastModifiedBy], 'Ingo Klemm'
    assert_nil   metadata[:core][:revision]
    assert_equal metadata[:core][:category], "Test"
    assert_equal metadata[:core][:created], '2015-06-05T18:19:34Z'
    assert_equal metadata[:core][:modified], '2024-12-15T10:36:33Z'
    assert_equal metadata[:core][:application], 'Microsoft Excel'
    assert_equal metadata[:core][:company], ''

    assert_equal metadata[:custom]["Pfad"], "\"C:\\Users\\iklemm\\Documents\\Manuals\\Vue Handbook.pdf\""
    assert_equal metadata[:custom]["Zweck"], "TISAX"
  end
end
