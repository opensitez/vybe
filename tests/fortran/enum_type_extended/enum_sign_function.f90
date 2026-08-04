! vybe-test: fortran/enum_type_extended/enum_sign_function
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: POS = 5, NEG = -5
end enum
if ((sign(POS, NEG)) /= -5) then
    print *, "FAIL: want [-5] got [", sign(POS, NEG), "]"
    stop 1
end if
end program t
