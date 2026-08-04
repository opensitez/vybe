! vybe-test: fortran/associate_construct_extended/associate_in_internal_subroutine
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
call show()
contains
subroutine show()
integer :: k = 11
associate (alias => k)
if ((alias) /= 11) then
    print *, "FAIL: want [11] got [", alias, "]"
    stop 1
end if
end associate
end subroutine show
end program t
