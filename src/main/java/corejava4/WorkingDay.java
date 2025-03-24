package corejava4;

import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

@Getter
@Setter
public class WorkingDay {
    private String date;
    private double totalHours;
    private double totalAmount;
    private List<Shift> shifts;

    public WorkingDay(String date) {
        this.date = date;
        this.totalHours = 0;
        this.totalAmount = 0;
        this.shifts = new ArrayList<>();
    }

    public void addShift(Shift shift, double hours, double amount) {
        shifts.add(shift);
        totalHours += shift.getHours();
        totalAmount += amount;
    }

    public String toString() {
        return "WorkingDay{" +
                "date='" + date + '\'' +
                ", totalHours=" + totalHours +
                ", totalAmount=" + totalAmount +
                ", shifts=" + shifts +
                '}';
    }
}
