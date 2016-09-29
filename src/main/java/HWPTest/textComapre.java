package HWPTest;

import java.io.File;
import java.io.IOException;
import java.util.Scanner;

public class textComapre
{
	public static void main(String[] args) throws IOException
	{
		Scanner keyboard = new Scanner(System.in);

		System.out.print("Enter first file's name: ");
		String file1 = keyboard.next();

		System.out.print("Enter second file's name: ");
		String file2 = keyboard.next();

		System.out.println(" ");

		File inputFile1 = new File(file1);
		Scanner infile1 = new Scanner(inputFile1);

		File inputFile2 = new File(file2);
		Scanner infile2 = new Scanner(inputFile2);

		compareFiles(infile1,infile2);
	}

		public static void compareFiles(Scanner infile1, Scanner infile2) throws IOException
		{
				int counter = 1;

				String line1 = readFrom(infile1);
				String line2 = readFrom(infile2);

				while(line1 != null && line2 != null)
				{
					int answer = line1.compareTo(line2);
					if(answer != 0)
					{
						System.out.println("Difference found in line " + counter);
						printLine("<",line1);
						printLine(">",line2);
					}

						if(infile1.hasNext() && infile2.hasNext())
						{
							line1 = readFrom(infile1);
							line2 = readFrom(infile2);
							counter++;
						}
						else
						break;

				}
		}
		
		public static String readFrom(Scanner infile) throws IOException
        {
           if(infile.hasNextLine())
             return infile.nextLine();
           return null;
        }


		public static void printLine(String prefix, String line)
		{
			System.out.println(prefix + line);
		}
	}