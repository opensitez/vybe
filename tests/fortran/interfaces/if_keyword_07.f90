! vybe-test: fortran/interfaces/if_keyword_07
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(x,y)
integer::x,y
end
program p
call s(y=2,x=1)
end program p
