! vybe-test: fortran/keyword_arguments/kw_08
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(i,j)
integer::i,j
end
program p
call s(j=2,i=1)
end program p
