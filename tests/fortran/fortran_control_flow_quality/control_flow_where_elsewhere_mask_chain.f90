! vybe-test: fortran/fortran_control_flow_quality/control_flow_where_elsewhere_mask_chain
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_where_elsewhere_mask_chain
    integer :: a(3)
    integer :: b(3)
    a = [1, -2, 3]
    b = 0
    where (a > 0)
        b = 10
    elsewhere (a < 0)
        b = 20
    elsewhere
        b = 30
    end where
    if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 20) then
    print *, "FAIL: want [20] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 10) then
    print *, "FAIL: want [10] got [", b(3), "]"
    stop 1
end if
end program control_flow_where_elsewhere_mask_chain
