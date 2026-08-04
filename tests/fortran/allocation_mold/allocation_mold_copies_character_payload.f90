! vybe-test: fortran/allocation_mold/allocation_mold_copies_character_payload
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program t
character(len=:), allocatable :: a, b
allocate(character(len=5) :: b)
b = 'moldx'
allocate(a, mold=b)
if ((len(a)) /= 5) then
    print *, "FAIL: want [5] got [", len(a), "]"
    stop 1
end if
if (trim(a) /= "moldx") then
    print *, "FAIL: want [moldx] got [", a, "]"
    stop 1
end if
end program t
