! vybe-test: fortran/enumerations/enum_three_03
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program t
enum, bind(c)
enumerator :: a=1, b=2, c=3
end enum
if (a /= 1) then
    print *, "FAIL: want [1] got [", a, "]"
    stop 1
end if
if (b /= 2) then
    print *, "FAIL: want [2] got [", b, "]"
    stop 1
end if
if (c /= 3) then
    print *, "FAIL: want [3] got [", c, "]"
    stop 1
end if
end program t
