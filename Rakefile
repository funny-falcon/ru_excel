require 'rake'
require 'rake/testtask'
require 'rake/rdoctask'

#desc 'Generate documentation for the ru_excel plugin.'
#Rake::RDocTask.new(:rdoc) do |rdoc|
#  rdoc.rdoc_dir = 'rdoc'
#  rdoc.title    = 'IdentityMap'
#  rdoc.options << '--line-numbers' << '--inline-source'
#  rdoc.rdoc_files.include('README')
#  rdoc.rdoc_files.include('lib/**/*.rb')
#end

begin
  require 'jeweler'
  Jeweler::Tasks.new do |gemspec|
    gemspec.name = "ru_excel"
    gemspec.summary = "Fast writting of MsExcel files (port of pyExcelerator)"
    gemspec.description = "Port of pyExcelerator tunned for faster .xls generation"
    gemspec.email = "funny.falcon@gmail.com"
    gemspec.homepage = "http://github.com/funny-falcon/ru_excel"
    gemspec.authors = ["Sokolov Yura aka funny_falcon"]
    #gemspec.add_dependency('')
    gemspec.rubyforge_project = 'ru-excel'
  end
  Jeweler::GemcutterTasks.new
  Jeweler::RubyforgeTasks.new
rescue LoadError
  puts "Jeweler not available. Install it with: gem install jeweler"
end
