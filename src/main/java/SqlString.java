package src.main.java;

import java.util.StringTokenizer;

public class SqlString {
    /**
     * Creates a SQL-compatible string from an input String object.  Any single
     * quote characters are replaced with two single quotes.
     * @param  s - String input.
     * @return String that can be used in SQL queries.
     */
    static public String encode(String s) {
        StringBuffer sb = new StringBuffer("");
        if (s == null)  {
            return (sb.toString());
        }

        try {
            // Look for single quote characters.
            StringTokenizer tok = new StringTokenizer(s, "'", true);
            while (tok.hasMoreTokens()) {
                // Append it to the buffer.
                String strTok = tok.nextToken();
                sb.append(strTok);

                // Add an extra quote after any quote characters found.
                if (strTok.compareTo("'") == 0) {
                    sb.append('\'');
                }
            }
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        return (sb.toString());
    }
}
