! vybe-test: fortran/associate_construct_extended/associate_ior_bitwise
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a = 5, b = 3
associate (bits => ior(a, b))
if ((bits) /= 7) then
    print *, "FAIL: want [7] got [", bits, "]"
    stop 1
end if
end associate
end program t
