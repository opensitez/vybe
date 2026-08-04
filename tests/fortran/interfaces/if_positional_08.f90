! vybe-test: fortran/interfaces/if_positional_08
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(x,y)
integer::x,y
end
program p
call s(1,2)
end program p
