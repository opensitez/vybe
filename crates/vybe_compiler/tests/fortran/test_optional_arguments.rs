use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(optional_arguments_01,"subroutine s(x)
integer, optional :: x
end subroutine s
");
c!(optional_arguments_02,"subroutine s(x,y)
integer, optional :: x,y
end subroutine s
");
c!(optional_arguments_03,"subroutine s(x)
real, optional :: x
end subroutine s
");
c!(optional_arguments_04,"subroutine s(x)
character(len=*), optional :: x
end subroutine s
");
c!(optional_arguments_05,"program p
interface
subroutine s(x)
integer, optional :: x
end subroutine s
end interface
call s()
end program p
");
c!(optional_arguments_06,"program p
interface
subroutine s(x)
integer, optional :: x
end subroutine s
end interface
call s(1)
end program p
");
c!(optional_arguments_07,"program p
interface
subroutine s(x,y)
integer, optional :: x,y
end subroutine s
end interface
call s(y=2)
end program p
");
c!(optional_arguments_08,"program p
interface
subroutine s(x,y)
integer, optional :: x,y
end subroutine s
end interface
call s(x=1)
end program p
");
c!(optional_arguments_09,"subroutine s(x)
integer, optional, value :: x
end subroutine s
");
c!(optional_arguments_10,"subroutine s(x)
integer, optional, intent(in) :: x
end subroutine s
");