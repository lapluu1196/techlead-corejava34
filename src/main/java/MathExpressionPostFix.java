import java.util.Scanner;
import java.util.Stack;

public class MathExpressionPostFix {
    private static int priority(char op) {
        switch (op) {
            case '+':
            case '-':
                return 1;
            case '*':
            case '/':
                return 2;
            default:
                return 0;
        }
    }

    public static String infixToPostfix(String expression) {
        Stack<Character> operators = new Stack<>();
        StringBuilder postfix = new StringBuilder();

        for (int i = 0; i < expression.length(); i++) {
            char c = expression.charAt(i);
            
            if (c == ' ') {
                continue;
            }
            if (Character.isDigit(c)) {
                postfix.append(c);
                while (i + 1 < expression.length() && Character.isDigit(expression.charAt(i + 1))) {
                    postfix.append(expression.charAt(++i));
                }
                postfix.append(' ');
            }
            else if (c == '(') {
                operators.push(c);
            }
            else if (c == ')') {
                while (!operators.isEmpty() && operators.peek() != '(') {
                    postfix.append(operators.pop()).append(' ');
                }
                operators.pop();
            }
            else if (c == '+' || c == '-' || c == '*' || c == '/') {
                while (!operators.isEmpty() && priority(c) <= priority(operators.peek())) {
                    postfix.append(operators.pop()).append(' ');
                }
                operators.push(c);
            }
        }

        while (!operators.isEmpty()) {
            postfix.append(operators.pop()).append(' ');
        }

        return postfix.toString();
    }

    public static double evaluatePostfix(String postfix) {
        Stack<Double> stack = new Stack<>();
        Scanner sc = new Scanner(postfix);

        while (sc.hasNext()) {
            String token = sc.next();
            if (token.matches("\\d+")) {
                stack.push(Double.parseDouble(token));
            } else {
                double b = stack.pop();
                double a = stack.pop();
                switch (token.charAt(0)) {
                    case '+':
                        stack.push(a + b);
                        break;
                    case '-':
                        stack.push(a - b);
                        break;
                    case '*':
                        stack.push(a * b);
                        break;
                    case '/':
                        if (b == 0) {
                            throw new ArithmeticException("Failed to divide by 0");
                        }
                        stack.push(a / b);
                        break;
                }
            }
        }

        return stack.pop();
    }
}
