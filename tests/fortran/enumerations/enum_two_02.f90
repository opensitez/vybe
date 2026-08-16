! vybe-test: fortran/enumerations/enum_two_02
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program t
enum, bind(c)
enumerator :: red=1, green=2
end enum
if (red /= 1) then
    print *, "FAIL: want [1] got [", red, "]"
    stop 1
end if
if (green /= 2) then
    print *, "FAIL: want [2] got [", green, "]"
    stop 1
end if
end program t
