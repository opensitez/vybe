! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_vector_alias_preserved_after_reassign
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_vector_alias_preserved_after_reassign
    integer :: a(1:6)
    integer :: b(1:6)
    a = (/1, 2, 3, 4, 5, 6/)
    b = a
    a(2:5:2) = b(1:2)
    if ((a(2)) /= 1) then
    print *, "FAIL: want [1] got [", a(2), "]"
    stop 1
end if
    if ((a(4)) /= 2) then
    print *, "FAIL: want [2] got [", a(4), "]"
    stop 1
end if
    if ((sum(a)) /= 21) then
    print *, "FAIL: want [21] got [", sum(a), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_vector_alias_preserved_after_reassign
