! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_zero_sized_allocation
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_zero_sized_allocation
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 0, 0 ]
    integer, allocatable :: values(:)
    integer :: expected_size
    allocate(values(0))
    values = 4
    expected_size = size(values)
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((expected_size) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", expected_size, "]"
        stop 1
    end if
    if (expected_size == 0) print *, sum(values)
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program array_fill_pattern_internals_fill_zero_sized_allocation
