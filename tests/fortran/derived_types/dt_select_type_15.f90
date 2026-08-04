! vybe-test: fortran/derived_types/dt_select_type_15
! origin: languages/fortran/tests/fortran/test_derived_types.rs
program p
class(*), allocatable :: x
allocate(integer :: x)
select type(x)
 type is(integer)
  print *, x
 class default
end select
end program p
