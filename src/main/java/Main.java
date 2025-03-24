import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        try (Scanner scanner = new Scanner(System.in)) {
            System.out.println("Please enter expression: ");
            String expression = scanner.nextLine();

            try {
                String postfix = MathExpressionPostFix.infixToPostfix(expression);
                double result = MathExpressionPostFix.evaluatePostfix(postfix);
                System.out.println("Postfix: " + postfix);
                System.out.println("Result: " + result);
            } catch (Exception e) {
                System.out.println("Error: " + e.getMessage());
            }
        }
    }
}
