! vybe-test: fortran/associate_construct_extended/associate_inside_if_then
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: n = 6
if (n > 0) then
associate (sq => n * n)
if ((sq) /= 36) then
    print *, "FAIL: want [36] got [", sq, "]"
    stop 1
end if
end associate
end if
end program t
