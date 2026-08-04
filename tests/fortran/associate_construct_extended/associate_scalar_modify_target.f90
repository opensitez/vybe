! vybe-test: fortran/associate_construct_extended/associate_scalar_modify_target
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: count = 1
associate (c => count)
c = c + 4
end associate
if ((count) /= 5) then
    print *, "FAIL: want [5] got [", count, "]"
    stop 1
end if
end program t
