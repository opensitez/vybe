! vybe-test: fortran/enumerations/enum_large_08
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program t
enum, bind(c)
enumerator :: big=1000
end enum
if (big /= 1000) then
    print *, "FAIL: want [1000] got [", big, "]"
    stop 1
end if
end program t
