! vybe-test: fortran/attributes/attr_dimension_declared_array_behaves_like_fixed_shape
! origin: languages/fortran/tests/fortran/test_attributes.rs

program attr_dimension_declared_array_behaves_like_fixed_shape
    integer, dimension(3) :: a
    a = [5, 6, 7]
    if ((a(1) + a(2) + a(3)) /= 18) then
    print *, "FAIL: want [18] got [", a(1) + a(2) + a(3), "]"
    stop 1
end if
end program attr_dimension_declared_array_behaves_like_fixed_shape
