(() => {
    const view = document.querySelector('.calendar-day-view, .calendar-week-view');
    if (!view) return;

    const isDayView = view.classList.contains('calendar-day-view');
    const rows = [...view.querySelectorAll(isDayView ? '.calendar-day-slot' : '.calendar-hour-row')];
    const eventSelector = '.calendar-event-link, .calendar-scheduled-event';
    const now = new Date();
    const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    const isCurrentDay = isDayView
        ? document.getElementById('calendar-date')?.value === today
        : Boolean(view.querySelector('.calendar-weekday[aria-current="date"]'));
    let target;

    if (isCurrentDay) {
        if (isDayView) {
            const currentMinute = now.getMinutes() < 30 ? 0 : 30;
            target = rows.find(row => Number(row.dataset.calendarHour) === now.getHours()
                && Number(row.dataset.calendarMinute) === currentMinute);
        } else {
            target = rows.find(row => Number(row.dataset.calendarHour) === now.getHours());
        }

        if (target) {
            const indicator = document.createElement('span');
            indicator.className = 'calendar-now-indicator';
            indicator.setAttribute('aria-hidden', 'true');
            const minuteInSlot = isDayView ? now.getMinutes() % 30 : now.getMinutes();
            const slotLength = isDayView ? 30 : 60;
            indicator.style.setProperty('--calendar-now-offset', `${minuteInSlot / slotLength * 100}%`);
            const indicatorContainer = isDayView
                ? target.querySelector('.calendar-day-slot-events')
                : [...target.children].find(cell => cell.classList.contains('calendar-day-today'));
            indicatorContainer?.append(indicator);
        }
    }

    target ??= rows.find(row => row.querySelector(eventSelector));

    if (target) {
        requestAnimationFrame(() => {
            const top = target.getBoundingClientRect().top + window.scrollY - 100;
            window.scrollTo({ top: Math.max(0, top), behavior: 'instant' });
        });
    }
})();
