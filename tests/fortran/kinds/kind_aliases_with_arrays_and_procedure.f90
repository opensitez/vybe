! vybe-test: fortran/kinds/kind_aliases_with_arrays_and_procedure
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer, parameter :: ki = selected_int_kind(9)
    real, parameter :: kr = selected_real_kind(6, 37)
    integer(kind=ki), dimension(4) :: a = [1, 2, 3, 4]
    real(kind=kr), dimension(3) :: b = [1.0, 2.0, 3.0]
    if ((size(a)) /= 4) then
    print *, "FAIL: want [4] got [", size(a), "]"
    stop 1
end if
    if ((size(b)) /= 3) then
    print *, "FAIL: want [3] got [", size(b), "]"
    stop 1
end if
    if ((kind(a)) /= 8) then
    print *, "FAIL: want [8] got [", kind(a), "]"
    stop 1
end if
    if ((kind(b)) /= 8) then
    print *, "FAIL: want [8] got [", kind(b), "]"
    stop 1
end if
    if ((sum(a)) /= 10) then
    print *, "FAIL: want [10] got [", sum(a), "]"
    stop 1
end if
end program test
