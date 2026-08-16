! vybe-test: fortran/enumerations/enum_negative_07
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program t
enum, bind(c)
enumerator :: a=-1, b=0
end enum
if (a /= -1) then
    print *, "FAIL: want [-1] got [", a, "]"
    stop 1
end if
if (b /= 0) then
    print *, "FAIL: want [0] got [", b, "]"
    stop 1
end if
end program t
