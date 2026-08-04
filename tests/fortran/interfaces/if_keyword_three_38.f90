! vybe-test: fortran/interfaces/if_keyword_three_38
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(z=3,x=1,y=2)
end program p
