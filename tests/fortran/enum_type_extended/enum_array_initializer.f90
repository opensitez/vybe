! vybe-test: fortran/enum_type_extended/enum_array_initializer
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: R = 0, G = 1, B = 2
end enum
integer :: codes(3) = [R, G, B]
if ((codes(2)) /= 1) then
    print *, "FAIL: want [1] got [", codes(2), "]"
    stop 1
end if
end program t
