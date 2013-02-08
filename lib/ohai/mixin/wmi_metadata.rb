#
# Author:: Larry Gilbert (<larry@L2G.to>)
# Copyright:: Copyright (c) 2013 Opscode, Inc.
# License:: Apache License, Version 2.0
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

gem 'ruby-wmi', '~> 0.4.0'
require 'ruby-wmi'

module Ohai
  module Mixin
    module WmiMetadata

      # Extract properties from a RubyWMI object.  Since an attribute name may
      # be reported that Windows can't look up, catch this error and handle it
      # gracefully by silently ignoring it.
      # @param wmi_object [RubyWMI] WMI object
      # @return [Mash]

      def extract_wmi_properties_to_mash(wmi_object)
        mash = Mash.new
        wmi_object.attribute_names.each do |prop_name|
          value = get_wmi_property(wmi_object, prop_name)
          mash[prop_name.to_sym] = value unless value.nil?
        end

        mash
      end

      # Safely get a property from a WMI object, returning nil if there is
      # no such property.
      # @param wmi_object [RubyWMI] WMI object
      # @param property_name [String] property to look up on the object

      def get_wmi_property(wmi_object, property_name)
        begin
          wmi_object.send(property_name)
        rescue WIN32OLERuntimeError
          nil
        end
      end

    end
  end
end

# vim:ts=2:sw=2:
