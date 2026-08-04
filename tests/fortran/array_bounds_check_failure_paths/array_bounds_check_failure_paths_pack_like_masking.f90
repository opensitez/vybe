! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_pack_like_masking
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_pack_like_masking
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    integer :: a(1:6)
    integer :: i
    integer :: kept

    a = (/ 3, -1, 4, -2, 5, -3 /)
    kept = 0
    do i = 1, 6
        if (i >= lbound(a, 1) .and. i <= ubound(a, 1) .and. a(i) > 0) then
            kept = kept + 1
        end if
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((kept) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", kept, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program array_bounds_check_failure_paths_pack_like_masking
