export function parseDate(dateInput: string): Date | null {
  if (!dateInput) return null;

  // Handle relative dates — all in UTC to match Planner's dueDateTime storage
  const now = new Date();
  const today = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));

  switch (dateInput.toLowerCase()) {
    case "today":
      return today;
    case "tomorrow":
      return new Date(today.getTime() + 24 * 60 * 60 * 1000);
    case "yesterday":
      return new Date(today.getTime() - 24 * 60 * 60 * 1000);
    case "this-week": {
      const startOfWeek = new Date(today);
      startOfWeek.setUTCDate(today.getUTCDate() - today.getUTCDay()); // Sunday UTC
      return startOfWeek;
    }
    case "next-week": {
      const startOfNextWeek = new Date(today);
      startOfNextWeek.setUTCDate(today.getUTCDate() + (7 - today.getUTCDay())); // Next Sunday UTC
      return startOfNextWeek;
    }
    case "last-week": {
      const startOfLastWeek = new Date(today);
      startOfLastWeek.setUTCDate(today.getUTCDate() - today.getUTCDay() - 7); // Last Sunday UTC
      return startOfLastWeek;
    }
    default:
      // Try parsing as ISO 8601 or other date formats
      const parsed = new Date(dateInput);
      return isNaN(parsed.getTime()) ? null : parsed;
  }
}

export function isDateInRange(dateString: string | null, startDate: Date, endDate: Date): boolean {
  if (!dateString) return false;
  const date = new Date(dateString);
  if (isNaN(date.getTime())) return false;
  return date >= startDate && date <= endDate;
}
