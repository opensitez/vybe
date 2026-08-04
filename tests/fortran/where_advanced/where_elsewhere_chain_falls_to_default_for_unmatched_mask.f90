! vybe-test: fortran/where_advanced/where_elsewhere_chain_falls_to_default_for_unmatched_mask
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(4) = [5, 15, 25, 35]
    character(len=6) :: b(4)
    where (a < 10)
        b = 'low'
    elsewhere (a < 20)
        b = 'mid'
    elsewhere (a < 30)
        b = 'high'
    elsewhere
        b = 'top'
    end where
    if (trim(trim(b(1))) /= "low") then
    print *, "FAIL: want [low] got [", trim(b(1)), "]"
    stop 1
end if
    if (trim(trim(b(2))) /= "mid") then
    print *, "FAIL: want [mid] got [", trim(b(2)), "]"
    stop 1
end if
    if (trim(trim(b(3))) /= "high") then
    print *, "FAIL: want [high] got [", trim(b(3)), "]"
    stop 1
end if
    if (trim(trim(b(4))) /= "top") then
    print *, "FAIL: want [top] got [", trim(b(4)), "]"
    stop 1
end if
end program test
