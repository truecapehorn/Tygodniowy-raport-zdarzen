(() => {
    const view = document.querySelector('.calendar-day-view, .calendar-week-view');
    if (!view) return;

    const isDayView = view.classList.contains('calendar-day-view');
    const rows = [...view.querySelectorAll(isDayView ? '.calendar-day-slot' : '.calendar-hour-row')];
    const eventSelector = '.calendar-event-link, .calendar-scheduled-event';
    let target = rows.find(row => row.querySelector(eventSelector));

    if (!target) {
        const now = new Date();
        const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
        const shouldUseCurrentHour = isDayView
            ? document.getElementById('calendar-date')?.value === today
            : Boolean(view.querySelector('.calendar-weekday[aria-current="date"]'));
        if (shouldUseCurrentHour) {
            target = rows.find(row => Number(row.dataset.calendarHour) === now.getHours());
        }
    }

    if (target) {
        requestAnimationFrame(() => {
            const top = target.getBoundingClientRect().top + window.scrollY - 100;
            window.scrollTo({ top: Math.max(0, top), behavior: 'instant' });
        });
    }
})();
