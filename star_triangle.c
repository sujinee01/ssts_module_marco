#include <stdio.h>

int main(void)
{
    int i = 0;
    
    for (; i<6; i++) {
        for(int j=0; j=<i; j++) {
            printf("%c", '*');
        }
        printf("/n");
    }


    return 0;
}