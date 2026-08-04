! vybe-test: fortran/allocation_mold/allocation_mold_04
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program p
character(len=:), allocatable :: s, t
allocate(character(len=4) :: t)
allocate(s, mold=t)
end program p
