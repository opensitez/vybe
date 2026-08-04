! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_string_slice_like_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_string_slice_like_bounds
    character(len=8) :: word
    integer :: idx
    integer :: status

    word = 'fortran '
    idx = 0
    if (idx >= 1 .and. idx <= len(word)) then
        status = ichar(word(idx:idx))
    else
        status = -1
    end if
    if ((status) /= -1) then
    print *, "FAIL: want [-1] got [", status, "]"
    stop 1
end if

    idx = 8
    if (idx >= 1 .and. idx <= len(word)) then
        status = ichar(word(idx:idx))
    else
        status = -1
    end if
    if ((status) /= 32) then
    print *, "FAIL: want [32] got [", status, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_string_slice_like_bounds
