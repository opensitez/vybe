! vybe-test: fortran/enumerations/enum_bindc_01
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program t
enum, bind(c)
enumerator :: red
end enum
if (red /= 0) then
    print *, "FAIL: want [0] got [", red, "]"
    stop 1
end if
end program t
