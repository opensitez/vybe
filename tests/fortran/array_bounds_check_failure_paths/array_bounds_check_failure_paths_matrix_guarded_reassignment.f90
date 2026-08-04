! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_matrix_guarded_reassignment
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_matrix_guarded_reassignment
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 1, -1 ]
    integer :: src(1:2, 1:2)
    integer :: dst(1:2, 1:2)
    integer :: i
    integer :: j

    src = reshape((/ 1, 2, 3, 4 /), (/2,2/))
    dst = 0

    do i = 0, 2
        do j = 0, 2
            if (i >= lbound(src, 1) .and. i <= ubound(src, 1) .and. &
                j >= lbound(src, 2) .and. j <= ubound(src, 2)) then
                dst(i, j) = src(i, j)
            else
                dst(i, j) = -1
            end if
        end do
    end do

        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((dst(1,1)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", dst(1,1), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((dst(0,0)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", dst(0,0), "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program array_bounds_check_failure_paths_matrix_guarded_reassignment
