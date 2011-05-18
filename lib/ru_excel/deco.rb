class TrueClass
  def to_i; 1; end
end
class FalseClass
  def to_i; 0; end
end

module Excel
  module Deco
    def accepts(method, *types)
	class_eval do
	    method_accepts = "#{method}_accepts".to_sym
	    alias_method method_accepts, method.to_sym
	    define_method(method.to_sym) do |*args|
		for t,a in types.zip(args)
		    raise "Type mismatch" unless t === a
		end
		self.send(method_accepts, *args)
	    end
	end
    end

    def accepts_bool(*methods)
      methods.each{|method|
	class_eval do
	    method_accepts = "#{method}_accepts_bool".to_sym
	    alias_method method_accepts, method.to_sym
	    define_method(method.to_sym) do |bool|
		raise "Type mismatch" if bool!=true and bool != false
		self.send(method_accepts, bool)
	    end
	end
      }
    end
	    
    def returns(method, *types)
	class_eval do
	    method_returns = "#{method}_returns".to_sym
	    alias_method method_returns, method.to_sym
	    define_method(method.to_sym) do |*args|
		ret = self.send(method_returns, *args)
		for t,r in types.zip(ret)
		    raise "Type ret mismatch" unless t===a
		end
		ret
	    end
	end
    end
    
    def short_accessor(*attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    #raise "Type Mismatch" if not Integer === value
		    @#{attr} = value.to_i & 0xFFFF
		end
                attr_reader :#{attr}
	    EOF;
	end
    end

    def int_accessor(*attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    #raise "Type Mismatch" if not Integer === value
		    @#{attr} = value.to_i
		end
                attr_reader :#{attr}
	    EOF;
	end
    end

    def absint_accessor(*attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    #raise "Type Mismatch" if not Integer === value
		    @#{attr} = value.to_i.abs
		end
                attr_reader :#{attr}
	    EOF;
	end
    end
    
    def bool_int_accessor(*attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    #raise "Type Mismatch" if value != true and value != false
		    @#{attr} = value.to_i
		end
		def #{attr}
		    [false, true][@#{attr}]
		end
	    EOF;
	end
    end

    def bool_accessor(*attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    @#{attr} = value ? true : false
		end
		attr_reader :#{attr}
	    EOF;
	end
    end

    def float_accessor(*attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    @#{attr} = value.to_f
		end
                attr_reader :#{attr}
	    EOF;
	end
    end

    def string_accessor(*attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    @#{attr} = value.to_s
		end
                attr_reader :#{attr}
	    EOF;
	end
    end

    def array_accessor(*attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    @#{attr} = value.to_a
		end
                attr_reader :#{attr}
	    EOF;
	end
    end
    
    def type_accessor(type, *attrs)
	attrs.each do |attr|
	    module_eval <<-"EOF;"
		def #{attr}=(value)
		    raise "Type Mismatch" if not #{type.name} === value
		    @#{attr} = value
		end
                attr_reader :#{attr}
	    EOF;
	end
    end
  end
end