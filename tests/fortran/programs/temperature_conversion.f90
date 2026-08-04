! vybe-test: fortran/programs/temperature_conversion
! origin: languages/fortran/tests/fortran/test_programs.rs

program test
    real :: celsius, fahrenheit
    celsius = 100.0
    fahrenheit = celsius * 9.0 / 5.0 + 32.0
    if ((fahrenheit) /= 212) then
    print *, "FAIL: want [212] got [", fahrenheit, "]"
    stop 1
end if
end program test
