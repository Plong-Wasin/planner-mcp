export function parseDate(dateInput: string): Date | null {
  if (!dateInput) return null;

  // Handle relative dates
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  switch (dateInput.toLowerCase()) {
    case "today":
      return today;
    case "tomorrow":
      return new Date(today.getTime() + 24 * 60 * 60 * 1000);
    case "yesterday":
      return new Date(today.getTime() - 24 * 60 * 60 * 1000);
    case "this-week": {
      const startOfWeek = new Date(today);
      startOfWeek.setDate(today.getDate() - today.getDay()); // Sunday
      return startOfWeek;
    }
    case "next-week": {
      const startOfNextWeek = new Date(today);
      startOfNextWeek.setDate(today.getDate() + (7 - today.getDay())); // Next Sunday
      return startOfNextWeek;
    }
    case "last-week": {
      const startOfLastWeek = new Date(today);
      startOfLastWeek.setDate(today.getDate() - today.getDay() - 7); // Last Sunday
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
