! vybe-test: fortran/keyword_arguments/kw_22
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(name, value, scale)
character(len=*)::name
integer::value, scale
end
program p
call s(name='abc', value=3, scale=2)
end program p
