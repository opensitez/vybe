use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(nml_basic_01,"program p
integer::x=1
namelist /grp/ x
write(*,nml=grp)
end program p
");
c!(nml_array_02,"program p
integer::a(3)=[1,2,3]
namelist /grp/ a
write(*,nml=grp)
end program p
");
c!(nml_two_vars_03,"program p
integer::a=1,b=2
namelist /grp/ a,b
write(*,nml=grp)
end program p
");
c!(nml_real_04,"program p
real::x=1.5
namelist /grp/ x
write(*,nml=grp)
end program p
");
c!(nml_char_05,"program p
character(len=5)::s='abc'
namelist /grp/ s
write(*,nml=grp)
end program p
");
c!(nml_logical_06,"program p
logical::l=.true.
namelist /grp/ l
write(*,nml=grp)
end program p
");
c!(nml_complex_07,"program p
complex::z=(1.0,2.0)
namelist /grp/ z
write(*,nml=grp)
end program p
");
c!(nml_derived_08,"type::t
integer::x
end type t
program p
type(t)::v
namelist /grp/ v
write(*,nml=grp)
end program p
");
c!(nml_pointer_09,"program p
integer,target::x=1
integer,pointer::p
p=>x
namelist /grp/ p
write(*,nml=grp)
end program p
");
c!(nml_internal_10,"program p
integer::x=1
character(len=50)::buf
namelist /grp/ x
write(buf,nml=grp)
print *, trim(buf)
end program p
");
c!(nml_read_11,"program p
integer::x
character(len=50)::buf='&grp x=1 /'
namelist /grp/ x
read(buf,nml=grp)
print *, x
end program p
");
c!(nml_group2_12,"program p
integer::x=1
namelist /a/ x
write(*,nml=a)
end program p
");
c!(nml_group3_13,"program p
integer::x=1
namelist /numbers/ x
write(*,nml=numbers)
end program p
");
c!(nml_multi_arr_14,"program p
integer::a(2)=[1,2],b(2)=[3,4]
namelist /grp/ a,b
write(*,nml=grp)
end program p
");
c!(nml_nested_type_15,"type::t
integer::x
end type t
program p
type(t)::a(2)
namelist /grp/ a
write(*,nml=grp)
end program p
");
c!(nml_char_arr_16,"program p
character(len=3)::a(2)=(/'abc','def'/)
namelist /grp/ a
write(*,nml=grp)
end program p
");
c!(nml_default_vals_17,"program p
integer::x=0
namelist /grp/ x
write(*,nml=grp)
end program p
");
c!(nml_read_internal_18,"program p
integer::x=0
character(len=50)::buf='&grp x=7 /'
namelist /grp/ x
read(buf,nml=grp)
end program p
");
c!(nml_write_internal_19,"program p
integer::x=3
character(len=50)::buf
namelist /grp/ x
write(buf,nml=grp)
end program p
");
c!(nml_two_groups_20,"program p
integer::x=1,y=2
namelist /g1/ x
namelist /g2/ y
write(*,nml=g1)
write(*,nml=g2)
end program p
");