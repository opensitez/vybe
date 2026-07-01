use super::helpers::compile_ok;

macro_rules! c {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() { compile_ok($src); }
    };
}

c!(json_object_01, "program p
implicit none
character(len=64) :: s
s = '{\"a\":1}'
print *, s
end program p");
c!(json_array_02, "program p
implicit none
character(len=64) :: s
s = '[1,2,3]'
print *, s
end program p");
c!(json_nested_03, "program p
implicit none
character(len=128) :: s
s = '{\"a\":{\"b\":2}}'
print *, s
end program p");
c!(json_bool_04, "program p
implicit none
character(len=32) :: s
s = '{\"ok\":true}'
print *, s
end program p");
c!(json_null_05, "program p
implicit none
character(len=32) :: s
s = '{\"x\":null}'
print *, s
end program p");
c!(json_escape_06, "program p
implicit none
character(len=64) :: s
s = '{\"t\":\"a\\nb\"}'
print *, s
end program p");
c!(json_number_07, "program p
implicit none
character(len=64) :: s
s = '{\"n\":12.5}'
print *, s
end program p");
c!(json_whitespace_08, "program p
implicit none
character(len=64) :: s
s = '{ \"a\" : 1 }'
print *, s
end program p");
c!(json_chars_09, "program p
implicit none
character(len=64) :: s
s = '{\"name\":\"fortran\"}'
print *, s
end program p");
c!(json_longer_10, "program p
implicit none
character(len=128) :: s
s = '{\"items\":[{\"id\":1},{\"id\":2}]}'
print *, s
end program p");
c!(xml_basic_11, "program p
implicit none
character(len=64) :: s
s = '<a/>'
print *, s
end program p");
c!(xml_element_12, "program p
implicit none
character(len=64) :: s
s = '<a>1</a>'
print *, s
end program p");
c!(xml_nested_13, "program p
implicit none
character(len=128) :: s
s = '<a><b>2</b></a>'
print *, s
end program p");
c!(xml_attr_14, "program p
implicit none
character(len=64) :: s
s = '<a id=\"1\"/>'
print *, s
end program p");
c!(xml_text_15, "program p
implicit none
character(len=64) :: s
s = '<msg>hello</msg>'
print *, s
end program p");
c!(xml_cdata_16, "program p
implicit none
character(len=96) :: s
s = '<a><![CDATA[x]]></a>'
print *, s
end program p");
c!(xml_decl_17, "program p
implicit none
character(len=96) :: s
s = '<?xml version=\"1.0\"?><a/>'
print *, s
end program p");
c!(xml_namespace_18, "program p
implicit none
character(len=96) :: s
s = '<ns:a xmlns:ns=\"u\"/>'
print *, s
end program p");
c!(xml_generate_style_19, "program p
implicit none
character(len=64) :: s
s = '<row><id>1</id></row>'
print *, s
end program p");
c!(xml_json_mix_20, "program p
implicit none
character(len=160) :: a, b
a = '<row><id>1</id></row>'
b = '{\"id\":1}'
print *, a
print *, b
end program p");