! vybe-test: fortran/derived_type_copy_out_copy_in/test_derived_type_copy_out_copy_in_assigns_components
! origin: languages/fortran/tests/fortran/test_derived_type_copy_out_copy_in.rs

program test_derived_type_copy_out_copy_in
    type :: pair
        integer :: a
        integer :: b
    end type

    type(pair) :: source
    type(pair) :: dest

    source%a = 2
    source%b = 5
    dest = source

    if ((dest%a) /= 2) then
    print *, "FAIL: want [2] got [", dest%a, "]"
    stop 1
end if
    if ((dest%b) /= 5) then
    print *, "FAIL: want [5] got [", dest%b, "]"
    stop 1
end if
end program test_derived_type_copy_out_copy_in
