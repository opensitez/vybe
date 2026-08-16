! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_logical_repetition_through_merge
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program t
    logical, allocatable :: flags(:)
    integer :: n
    flags = (/ (.true., i = 1, 2), (.false., i = 1, 3) /)
    n = size(flags)
    if ((n) /= 5) then
    print *, "FAIL: want [5] got [", n, "]"
    stop 1
end if
    if ((count(flags)) /= 2) then
    print *, "FAIL: want [2] got [", count(flags), "]"
    stop 1
end if
    if ((merge(1, 0, flags(1))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, flags(1)), "]"
    stop 1
end if
    if ((merge(1, 0, flags(n))) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, flags(n)), "]"
    stop 1
end if
end program t
