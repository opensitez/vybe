! vybe-test: fortran/character_arrays/char_array_sort_bubble
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=5) :: arr(4) = ['delta', 'alpha', 'gamma', 'beta ']
    character(len=5) :: tmp
    integer :: i, j
    do i = 1, 3
        do j = 1, 4 - i
            if (arr(j) > arr(j+1)) then
                tmp = arr(j)
                arr(j) = arr(j+1)
                arr(j+1) = tmp
            end if
        end do
    end do
    print *, trim(arr(1))
    print *, trim(arr(4))
end program test
