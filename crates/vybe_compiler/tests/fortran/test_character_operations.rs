use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(chop_01,"program p
character(len=5) :: s='hello'
print *, s(1:2)
end program p
");
c!(chop_02,"program p
character(len=5) :: s='hello'
s(2:3)='ZZ'
print *, s
end program p
");
c!(chop_03,"program p
print *, 'a'//'b'
end program p
");
c!(chop_04,"program p
print *, repeat('x',4)
end program p
");
c!(chop_05,"program p
print *, trim('abc   ')
end program p
");
c!(chop_06,"program p
print *, adjustl('  abc')
end program p
");
c!(chop_07,"program p
print *, adjustr('abc  ')
end program p
");
c!(chop_08,"program p
print *, scan('abc123','0123456789')
end program p
");
c!(chop_09,"program p
print *, verify('abc','abc')
end program p
");
c!(chop_10,"program p
print *, index('fortran','tran')
end program p
");
c!(chop_11,"program p
print *, len_trim('ab   ')
end program p
");
c!(chop_12,"program p
print *, ichar('A')
end program p
");
c!(chop_13,"program p
print *, achar(65)
end program p
");
c!(chop_14,"program p
logical :: l
l = 'a' < 'b'
print *, l
end program p
");
c!(chop_15,"program p
logical :: l
l = 'A' /= 'a'
print *, l
end program p
");
c!(chop_16,"program p
character(len=6) :: s
s = 'ab'//'cd'//'ef'
print *, s
end program p
");
c!(chop_17,"program p
character(len=4) :: a(2)
a = ['ab  ','cd  ']
print *, a
end program p
");
c!(chop_18,"program p
character(len=20) :: buf
write(buf,'(A)') 'abc'
print *, trim(buf)
end program p
");
c!(chop_19,"program p
character(len=20) :: buf='42'
integer :: x
read(buf,*) x
print *, x
end program p
");
c!(chop_20,"program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
s='abc'
print *, s
end program p
");