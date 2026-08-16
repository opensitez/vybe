! vybe-test: fortran/enumerations/enum_case_14
! origin: languages/fortran/tests/fortran/test_enumerations.rs
program p
enum, bind(c)
enumerator :: a=1
end enum
select case(a)
case (1)
 print *,1
end select
end program p
