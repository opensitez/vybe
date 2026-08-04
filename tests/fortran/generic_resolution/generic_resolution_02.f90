! vybe-test: fortran/generic_resolution/generic_resolution_02
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
module m
interface g
module procedure s1,s2,s3
end interface
contains
subroutine s1(i)
integer::i
end
subroutine s2(r)
real::r
end
subroutine s3(c)
complex::c
end
end module m
