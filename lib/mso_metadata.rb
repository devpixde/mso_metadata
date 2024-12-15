# frozen_string_literal: true

require 'zip'
require 'nokogiri'

require_relative "mso_metadata/version"

module MsoMetadata
  class Error < StandardError; end
  def self.read office_xml_file
    core_metadata = {}
    custom_metadata = {}
    Zip::File.open(office_xml_file) do |zip_file|
      # Read the core metadata (docProps/core.xml)
      core_xml = zip_file.find_entry('docProps/core.xml')&.get_input_stream&.read
      if core_xml
        core_doc = Nokogiri::XML(core_xml)
        core_metadata[:title] = core_doc.at_xpath('//dc:title')&.text
        core_metadata[:subject] = core_doc.at_xpath('//dc:subject')&.text
        core_metadata[:creator] = core_doc.at_xpath('//dc:creator')&.text
        core_metadata[:keywords] = core_doc.at_xpath('//dc:keywords')&.text
        core_metadata[:description] = core_doc.at_xpath('//dc:description')&.text
        core_metadata[:lastModifiedBy] = core_doc.at_xpath('//cp:lastModifiedBy')&.text
        core_metadata[:revision] = core_doc.at_xpath('//cp:revision')&.text
        core_metadata[:category] = core_doc.at_xpath('//cp:category')&.text
        core_metadata[:created] = core_doc.at_xpath('//dcterms:created')&.text
        core_metadata[:modified] = core_doc.at_xpath('//dcterms:modified')&.text
      end

      # Read the app metadata (docProps/app.xml)
      app_xml = zip_file.find_entry('docProps/app.xml')&.get_input_stream&.read
      if app_xml
        app_doc = Nokogiri::XML(app_xml)
        core_metadata[:application] = app_doc.at_xpath('//*:Application')&.text
        core_metadata[:company] = app_doc.at_xpath('//*:Company')&.text
      end

      # Read the custom metadata (docProps/custom.xml)
      custom_xml = zip_file.find_entry('docProps/custom.xml')&.get_input_stream&.read
      if custom_xml
        custom_doc = Nokogiri::XML(custom_xml)
        custom_doc.xpath('//*:property').each do |property|
          value = property.text
          type_name = property.first_element_child.name
          # we convert bool and int, others are strings
          if type_name =~ /bool/
            value = value.downcase == 'true'
          elsif type_name =~ /i\d/
            value = value.to_i
          end
          custom_metadata[property.attr('name')] = value
        end
      end
    end
    { core: core_metadata, custom: custom_metadata }
  end
end
