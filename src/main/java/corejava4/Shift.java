package corejava4;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class Shift {
    private String shiftName;
    private double hours;
    private double amount;

    @Override
    public String toString() {
        return "Shift{" +
                "shiftName='" + shiftName + '\'' +
                ", hours=" + hours +
                ", amount=" + amount +
                '}';
    }
}
