! vybe-test: fortran/enumerations/enum_default_06
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program t
enum, bind(c)
enumerator :: a, b, c
end enum
if (a /= 0) then
    print *, "FAIL: want [0] got [", a, "]"
    stop 1
end if
if (b /= 1) then
    print *, "FAIL: want [1] got [", b, "]"
    stop 1
end if
if (c /= 2) then
    print *, "FAIL: want [2] got [", c, "]"
    stop 1
end if
end program t
