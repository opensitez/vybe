! vybe-test: fortran/types/array_declaration
! origin: languages/fortran/tests/fortran/test_types.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 30 ]
    integer, dimension(5) :: arr
    integer :: i
    do i = 1, 5
        arr(i) = i * 10
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((arr(3)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", arr(3), "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
