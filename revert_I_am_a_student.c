/*
 * revert_I_am_a_student.c: Revert a string such as "I am a student." to "student. a am I"
 * Copyright (C) 2015  Chih-Chyuan Hwang (hwangcc@csie.nctu.edu.tw)
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
 */

/*
 * 練習 編程之法 1.1 單詞反轉
 * Input: I am a student.
 * Output: student. a am I
 */

#include <stdio.h>
#include <string.h>

void revert_string(char *s, int from, int to)
{
    char temp;

    if (s) {
        while (from < to) {
            temp = s[from];
            s[from++] = s[to];
            s[to--] = temp;
        }
    }
}

void revert_I_am_a_student(char *string)
{
    char *current, *space;

    revert_string(string, 0, strlen(string) - 1);

    current = string;
    do {
        space = strchr(current, ' ');
        if (space) {
            revert_string(current, 0, space - current - 1);
            current = space + 1;
        } else {
            break;
        }
    } while (*current != '\0');
}

int main(int argc, char **argv)
{
    char string[4096];
    int n;

    printf("Input a string (max len 4095):");
    n = scanf("%[^\n]", string);
    if (n <= 0) {
        printf("No input string\nExit\n");
        return 0;
    } else {
        printf("Revert \"%s\"\n", string);
        revert_I_am_a_student(string);
        printf("Result: \"%s\"\n", string);
        return 1;
    }
}
