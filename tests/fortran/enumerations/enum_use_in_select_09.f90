! vybe-test: fortran/enumerations/enum_use_in_select_09
! origin: languages/fortran/tests/fortran/test_enumerations.rs
enum, bind(c)
enumerator :: a=1, b=2
end enum
program p
select case(a)
 case(1)
  print *,1
end select
end program p
