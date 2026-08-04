! vybe-test: fortran/arrays/array_2d_runtime_element_and_iteration
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 11, 32, 198 ]
    integer :: m(3,3)
    integer :: i, j, total
    total = 0
    do i = 1, 3
        do j = 1, 3
            m(i,j) = i * 10 + j
            total = total + m(i,j)
        end do
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((m(1,1)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", m(1,1), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((m(3,2)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", m(3,2), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program test
