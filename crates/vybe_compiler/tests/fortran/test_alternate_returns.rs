use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(alternate_returns_01,"subroutine s(*,*)
return 1
end
");
c!(alternate_returns_02,"subroutine s(*,*)
return 2
end
");
c!(alternate_returns_03,"subroutine s(*,*)
return
end
");
c!(alternate_returns_04,"program p
call s(*10,*20)
10 continue
20 continue
end program p
subroutine s(*,*)
return 1
end
");
c!(alternate_returns_05,"program p
call s(*10,*20)
10 continue
20 continue
end program p
subroutine s(*,*)
return 2
end
");
c!(alternate_returns_06,"program p
call s(*10,*20)
10 continue
20 continue
end program p
subroutine s(*,*)
return
end
");
c!(alternate_returns_07,"subroutine s(x,*,*)
integer::x
return 1
end
");
c!(alternate_returns_08,"subroutine s(x,*,*)
integer::x
return 2
end
");
c!(alternate_returns_09,"program p
integer::x=1
call s(x,*10,*20)
10 continue
20 continue
end program p
subroutine s(x,*,*)
integer::x
return 1
end
");
c!(alternate_returns_10,"program p
integer::x=2
call s(x,*10,*20)
10 continue
20 continue
end program p
subroutine s(x,*,*)
integer::x
return 2
end
");