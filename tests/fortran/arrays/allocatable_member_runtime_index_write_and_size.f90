! vybe-test: fortran/arrays/allocatable_member_runtime_index_write_and_size
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    type :: state
        integer, allocatable :: v(:)
    end type state
    type(state) :: value

    allocate(value%v(2))
    value%v(1) = 4
    value%v(2) = 6
    if ((value%v(1)) /= 4) then
    print *, "FAIL: want [4] got [", value%v(1), "]"
    stop 1
end if
    if ((value%v(2)) /= 6) then
    print *, "FAIL: want [6] got [", value%v(2), "]"
    stop 1
end if
    if ((size(value%v)) /= 2) then
    print *, "FAIL: want [2] got [", size(value%v), "]"
    stop 1
end if
end program test
