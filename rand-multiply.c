#include <stdio.h>
#include <time.h>
#include <stdlib.h>

int main(int argc, char **argv)
{
	int x, y;
	unsigned int seed;
	int ch;

	seed = (unsigned int)time(NULL);

	srand(seed);

	while (1) {
		x = random() % 1000;
		y = random() % 100;
		printf("x = %d, y = %d\n", x, y);

		printf("Press ENTER to get the answer\n");
		ch = getchar();
		if (ch < 0)
			printf("getchar() returned %d\n", ch);

		printf("x * y = %d\n", x * y);

		printf("Press ENTER to continue, or press 'Q' to exit\n");
		ch = getchar();
		if (ch < 0)
			printf("getchar() returned %d\n", ch);
		else if (ch == (int)'Q' || ch == (int)'q')
			break;
		else
			continue;
	}

	return 0;
}
