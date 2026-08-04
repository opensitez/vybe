! vybe-test: fortran/alternate_returns/alternate_returns_06
! origin: languages/fortran/tests/fortran/test_alternate_returns.rs
program p
call s(*10,*20)
10 continue
20 continue
end program p
subroutine s(*,*)
return
end
