# mso_metadata

## About
### Prerequisites

- Ruby 3.x

### Install

Add the following line to your application's Gemfile:

```ruby
gem 'mso_metadata'
```

And then execute:

```shell
bundle install
```

Or install it yourself as:

```shell
gem install mso_metadata
```

## Usage

### Reading

``` ruby
require 'mso_metadata'

# Create a Docx::Document object for our existing docx file
metadata = MsoMetadata.read('example.docx')

# metadata consists two hashes: metadata[:core] and metadata[:custom]
# custom metadata could be set with arbitrary key values. Thats why we have to split it, to prevent key value collisions.

puts metadata[:core]
puts metadata[:custom]

```

### Development

#### This was just a fast solution to wrap the code needed to retrieve the metadata from office files.

#### todo
- write some tests
