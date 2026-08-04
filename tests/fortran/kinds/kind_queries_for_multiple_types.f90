! vybe-test: fortran/kinds/kind_queries_for_multiple_types
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    logical :: l
    complex :: c
    character(len=5) :: s
    integer(kind=4) :: i
    real(kind=8) :: r
    l = .true.
    c = (1.0, 2.0)
    s = "abc"
    i = 5
    r = 3.0
    if ((kind(l)) /= 8) then
    print *, "FAIL: want [8] got [", kind(l), "]"
    stop 1
end if
    if ((kind(c)) /= 8) then
    print *, "FAIL: want [8] got [", kind(c), "]"
    stop 1
end if
    if ((kind(s)) /= 8) then
    print *, "FAIL: want [8] got [", kind(s), "]"
    stop 1
end if
    if ((kind(i)) /= 4) then
    print *, "FAIL: want [4] got [", kind(i), "]"
    stop 1
end if
    if ((kind(r)) /= 8) then
    print *, "FAIL: want [8] got [", kind(r), "]"
    stop 1
end if
end program test
