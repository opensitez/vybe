! vybe-test: fortran/arrays/assumed_size_parameter_array_runtime_is_one_based
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(4) = [ 5, 3, 8, 1 ]
    integer, parameter :: data(*) = [5, 3, 8, 1]
    integer :: i
    do i = 1, size(data)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 4) then
            print *, "FAIL: more than 4 line(s)"
            stop 1
        end if
        if ((data(i)) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", data(i), "]"
            stop 1
        end if
    end do
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program test
