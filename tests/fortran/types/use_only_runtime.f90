! vybe-test: fortran/types/use_only_runtime
! origin: languages/fortran/tests/fortran/test_types.rs

module mymod
    integer :: x = 10
    integer :: y = 20
end module mymod

program test
    use mymod, only: x
    if ((x) /= 10) then
    print *, "FAIL: want [10] got [", x, "]"
    stop 1
end if
end program test
