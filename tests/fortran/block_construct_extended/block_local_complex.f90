! vybe-test: fortran/block_construct_extended/block_local_complex
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
complex :: z
z = (3.0, 4.0)
if ((int(real(z) + aimag(z))) /= 7) then
    print *, "FAIL: want [7] got [", int(real(z) + aimag(z)), "]"
    stop 1
end if
end block
end program t
