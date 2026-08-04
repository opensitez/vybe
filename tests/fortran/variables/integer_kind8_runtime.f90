! vybe-test: fortran/variables/integer_kind8_runtime
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
integer(kind=8) :: k = 123456789012_8
if ((k) /= 123456789012) then
    print *, "FAIL: want [123456789012] got [", k, "]"
    stop 1
end if
end program t
