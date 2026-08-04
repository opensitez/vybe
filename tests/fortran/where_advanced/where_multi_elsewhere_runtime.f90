! vybe-test: fortran/where_advanced/where_multi_elsewhere_runtime
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(6) = [1, 5, 10, 50, 100, 500]
    character(len=6) :: b(6)
    where (a < 10)
        b = 'small '
    elsewhere (a < 100)
        b = 'medium'
    elsewhere
        b = 'large '
    end where
    if (trim(trim(b(1))) /= "small") then
    print *, "FAIL: want [small] got [", trim(b(1)), "]"
    stop 1
end if
    if (trim(trim(b(3))) /= "medium") then
    print *, "FAIL: want [medium] got [", trim(b(3)), "]"
    stop 1
end if
    if (trim(trim(b(5))) /= "large") then
    print *, "FAIL: want [large] got [", trim(b(5)), "]"
    stop 1
end if
end program test
