! vybe-test: fortran/keyword_arguments/kw_23
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(arr, n)
integer::arr(:)
integer::n
end
program p
call s(arr=[1,2,3], n=3)
end program p
