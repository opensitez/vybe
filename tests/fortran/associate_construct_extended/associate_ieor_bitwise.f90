! vybe-test: fortran/associate_construct_extended/associate_ieor_bitwise
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a = 5, b = 3
associate (bits => ieor(a, b))
if ((bits) /= 6) then
    print *, "FAIL: want [6] got [", bits, "]"
    stop 1
end if
end associate
end program t
