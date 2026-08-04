! vybe-test: fortran/types/array_declaration_runtime_indexed_assignment
! origin: languages/fortran/tests/fortran/test_types.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 10, 50 ]
    integer, dimension(5) :: arr
    integer :: i
    do i = 1, 5
        arr(i) = i * 10
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((arr(1)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", arr(1), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((arr(5)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", arr(5), "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test
