! vybe-test: fortran/alternate_returns/alternate_returns_10
! origin: languages/fortran/tests/fortran/test_alternate_returns.rs
program p
integer::x=2
call s(x,*10,*20)
10 continue
20 continue
end program p
subroutine s(x,*,*)
integer::x
return 2
end
