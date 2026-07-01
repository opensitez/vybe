use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(complex_kinds_01,"program p
complex(kind=4) :: z
print *, z
end program p
");
c!(complex_kinds_02,"program p
complex(kind=8) :: z
print *, z
end program p
");
c!(complex_kinds_03,"program p
complex :: z=(1.0,2.0)
print *, z
end program p
");
c!(complex_kinds_04,"program p
complex(kind=8) :: z=(1.0_8,2.0_8)
print *, z
end program p
");
c!(complex_kinds_05,"program p
complex :: a=(1.0,2.0), b=(3.0,4.0)
print *, a+b
end program p
");
c!(complex_kinds_06,"program p
complex :: a=(1.0,2.0)
print *, conjg(a)
end program p
");
c!(complex_kinds_07,"program p
complex :: a=(1.0,2.0)
print *, aimag(a)
end program p
");
c!(complex_kinds_08,"program p
print *, cmplx(1.0,2.0)
end program p
");
c!(complex_kinds_09,"program p
complex :: a=(1.0,2.0)
print *, real(a)
end program p
");
c!(complex_kinds_10,"program p
complex :: a=(1.0,2.0)
print *, abs(a)
end program p
");