! vybe-test: fortran/numerics/num_conv_04
! origin: languages/fortran/tests/fortran/test_numerics.rs
program p
integer::i
real::r=1.5
i=int(r)
print *,i
end program p
