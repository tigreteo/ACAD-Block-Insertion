namespace BlockInsertionV3
{
    class inchestoYards
    {
        //rounds to 1/8"
        public static Fraction makeFraction(double yardage)
        {
            string[] fraction = new string[2];
            int denom = 8;

            //round off the remaining yardage for pure yards
            //number can be changed later if rounds up
            int yards = (int)yardage / 36;

            //remove yards from total yards to leave inches
            double inches = yardage % 36;

            //round off remaining inches for the total inches
            int wholeInch = (int)inches;

            //remove the inches to get the fractional parts
            inches = inches - wholeInch;

            //divide by 1/8 to get numerator of 1/8ths
            inches = inches / (.125);

            //removed remainder to get whole numerator
            int numerator = (int)inches;

            //remove the whole num from the remainder
            double roundOff = inches - numerator;
            //return remainder to true length instead of portion of 1/8
            roundOff = roundOff * (.125);

            //use remainder to determine if numerator should be rounded up
            if (roundOff >= (.0625))
            { numerator++; }

            //use num/denom to deterimine if inches should round up, reduce to com denom ie 4/8 => 1/2
            if (numerator >= 8)
            { wholeInch++; }

            //use inches to determine if yds should round up
            if (inches >= 36)
            { yards++; }

            while ((numerator % 2 == 0 && denom % 2 == 0) == true)
            {
                numerator = numerator / 2;
                denom = denom / 2;
            }

            //place all into an array and return
            Fraction answer = new Fraction();
            answer.yds = yards;
            answer.whole = wholeInch;
            if (numerator != denom && numerator != 0)
            {
                answer.num = numerator;
                answer.denom = denom;
            }
            return answer;
        }
    }

    public struct Fraction
    {
        public int num, denom, yds;
        public double whole;
        public Fraction(int yards, int numerator, int denomerator, double inches)
        {
            yds = yards;
            whole = inches;
            num = numerator;
            denom = denomerator;
        }
    }
}
