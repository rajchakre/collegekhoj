/**
 * CollegeKhoj — Frontend Application Logic
 * Pure JavaScript with Debounced Search, Multi-Filter, Skeleton Loading, & Accessibility
 */

(() => {
  'use strict';

  // --- DOM Element References ---
  const searchInput     = document.getElementById('search-input');
  const searchClearBtn  = document.getElementById('search-clear-btn');
  const stateSelect     = document.getElementById('filter-state');
  const citySelect      = document.getElementById('filter-city');
  const courseSelect    = document.getElementById('filter-course');
  const resetBtn        = document.getElementById('reset-filters-btn');
  const activeTray      = document.getElementById('active-filters-tray');
  const activeBadgesList= document.getElementById('active-badges-list');
  const sortSelect      = document.getElementById('sort-select');
  const resultsGrid     = document.getElementById('colleges-grid');
  const resultsCount    = document.getElementById('results-count-badge');
  const navLiveCount    = document.getElementById('nav-live-count');

  // Stats Counters
  const statColleges    = document.getElementById('stat-colleges-count');
  const statCourses     = document.getElementById('stat-courses-count');
  const statStates      = document.getElementById('stat-states-count');

  // Modal Elements
  const modalBackdrop   = document.getElementById('college-modal-backdrop');
  const modalCloseBtn   = document.getElementById('modal-close-button');
  const modalContentSlot= document.getElementById('modal-content-slot');

  // --- State Variables ---
  let allColleges = [];
  let distinctStates = [];
  let distinctCities = [];
  let distinctCourses = [];

  const filterState = {
    query: '',
    state: '',
    city: '',
    course: '',
    sortBy: 'name-asc'
  };

  // --- Utility: Debounce ---
  function debounce(func, delay = 250) {
    let timeoutId;
    return (...args) => {
      clearTimeout(timeoutId);
      timeoutId = setTimeout(() => func(...args), delay);
    };
  }

  // --- Utility: Escape HTML for security ---
  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // --- Render Skeleton Loading Cards ---
  function renderSkeletons(count = 6) {
    let skeletonsHtml = '';
    for (let i = 0; i < count; i++) {
      skeletonsHtml += `
        <div class="skeleton-card" aria-hidden="true">
          <div class="skeleton-shimmer skeleton-badge"></div>
          <div class="skeleton-shimmer skeleton-title"></div>
          <div class="skeleton-shimmer skeleton-subtitle"></div>
          <div class="skeleton-shimmer skeleton-contact"></div>
          <div class="skeleton-tags">
            <div class="skeleton-shimmer skeleton-tag"></div>
            <div class="skeleton-shimmer skeleton-tag"></div>
            <div class="skeleton-shimmer skeleton-tag"></div>
          </div>
        </div>
      `;
    }
    resultsGrid.innerHTML = skeletonsHtml;
    resultsCount.textContent = 'Searching colleges…';
  }

  // Dynamic API base: adapts whether loaded on http://localhost:3000, Live Server, file:///, or Vercel
  const API_BASE = (window.location.protocol === 'file:' || (window.location.port !== '3000' && (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')))
    ? 'http://localhost:3000'
    : '';

  // --- Initialize App & Fetch Data ---
  async function initApp() {
    renderSkeletons(6);

    try {
      // Parallel fetch for colleges and courses to guarantee instant complete data
      const [collegesRes, coursesRes] = await Promise.all([
        fetch(`${API_BASE}/api/colleges`),
        fetch(`${API_BASE}/api/courses`)
      ]);

      if (!collegesRes.ok || !coursesRes.ok) {
        throw new Error('Failed to load data from server');
      }

      const collegesData = await collegesRes.json();
      const coursesData  = await coursesRes.json();

      if (!Array.isArray(collegesData) || !Array.isArray(coursesData)) {
        const msg = (collegesData && collegesData.error) || (coursesData && coursesData.error) || 'Unexpected response format';
        throw new Error(msg);
      }

      // Map courses to colleges
      const collegeCourseMap = {};
      coursesData.forEach(c => {
        if (!collegeCourseMap[c.college_id]) {
          collegeCourseMap[c.college_id] = [];
        }
        collegeCourseMap[c.college_id].push(c.course_name);
      });

      // Merge courses into colleges list
      allColleges = collegesData.map(col => ({
        ...col,
        courses: collegeCourseMap[col.id] || []
      }));

      // Extract unique values for filter dropdowns
      extractFilterOptions();
      populateDropdowns();
      updateMetricCounts(allColleges.length, coursesData.length, distinctStates.length);

      // Perform initial render
      applyFilters();

    } catch (err) {
      console.error(err);
      resultsGrid.innerHTML = `
        <div class="empty-state-card" role="alert">
          <div class="empty-state-icon-wrap" aria-hidden="true">⚠️</div>
          <h3 class="empty-state-title">Unable to Connect</h3>
          <p class="empty-state-desc">
            ${escapeHtml(err.message || 'Could not fetch colleges from the server.')}
          </p>
          <button type="button" class="btn-empty-reset" onclick="location.reload()">Retry Connection</button>
        </div>
      `;
      resultsCount.textContent = 'Connection error';
    }
  }

  // --- Extract Distinct Filters ---
  function extractFilterOptions() {
    const statesSet = new Set();
    const citiesSet = new Set();
    const coursesSet= new Set();

    allColleges.forEach(col => {
      if (col.state) statesSet.add(col.state.trim());
      if (col.city)  citiesSet.add(col.city.trim());
      if (Array.isArray(col.courses)) {
        col.courses.forEach(course => {
          if (course) coursesSet.add(course.trim());
        });
      }
    });

    distinctStates  = Array.from(statesSet).sort((a, b) => a.localeCompare(b));
    distinctCities  = Array.from(citiesSet).sort((a, b) => a.localeCompare(b));
    distinctCourses = Array.from(coursesSet).sort((a, b) => a.localeCompare(b));
  }

  // --- Populate Filter Selects ---
  function populateDropdowns() {
    // States
    stateSelect.innerHTML = '<option value="">All States</option>' +
      distinctStates.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');

    // Cities
    repopulateCityDropdown();

    // Courses
    courseSelect.innerHTML = '<option value="">All Courses</option>' +
      distinctCourses.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  }

  // Repopulate City dropdown (optionally filtered by selected state)
  function repopulateCityDropdown() {
    const selectedState = filterState.state;
    let filteredCities = [];

    if (selectedState) {
      const citySet = new Set();
      allColleges.forEach(col => {
        if (col.state === selectedState && col.city) {
          citySet.add(col.city.trim());
        }
      });
      filteredCities = Array.from(citySet).sort((a, b) => a.localeCompare(b));
    } else {
      filteredCities = distinctCities;
    }

    const currentCity = filterState.city;
    citySelect.innerHTML = '<option value="">All Cities</option>' +
      filteredCities.map(c => {
        const selected = (c === currentCity) ? ' selected' : '';
        return `<option value="${escapeHtml(c)}"${selected}>${escapeHtml(c)}</option>`;
      }).join('');
  }

  // --- Update Metric Overview ---
  function updateMetricCounts(collegesCount, coursesCount, statesCount) {
    if (statColleges) statColleges.textContent = collegesCount;
    if (statCourses)  statCourses.textContent = `${coursesCount}+`;
    if (statStates)   statStates.textContent = `${statesCount}`;
    if (navLiveCount) navLiveCount.textContent = `${collegesCount} Colleges across ${statesCount} States`;
  }

  // --- Filter and Sort Core Logic ---
  function applyFilters() {
    const q = filterState.query.trim().toLowerCase();
    const selectedState = filterState.state;
    const selectedCity  = filterState.city;
    const selectedCourse= filterState.course;

    const filtered = allColleges.filter(col => {
      // State match
      if (selectedState && col.state !== selectedState) {
        return false;
      }

      // City match
      if (selectedCity && col.city !== selectedCity) {
        return false;
      }

      // Course match
      if (selectedCourse) {
        const hasCourse = col.courses.some(c => c.toLowerCase() === selectedCourse.toLowerCase());
        if (!hasCourse) return false;
      }

      // Query search match (Name, City, State, or Courses)
      if (q) {
        const matchName   = col.name && col.name.toLowerCase().includes(q);
        const matchCity   = col.city && col.city.toLowerCase().includes(q);
        const matchState  = col.state && col.state.toLowerCase().includes(q);
        const matchCourse = col.courses && col.courses.some(c => c.toLowerCase().includes(q));

        if (!matchName && !matchCity && !matchState && !matchCourse) {
          return false;
        }
      }

      return true;
    });

    // Apply Sorting
    const sorted = sortColleges(filtered, filterState.sortBy);

    // Update active badges & Render
    updateActiveBadges();
    renderCollegeCards(sorted);
  }

  // --- Sort Helper ---
  function sortColleges(list, sortBy) {
    const copy = [...list];
    switch (sortBy) {
      case 'name-asc':
        return copy.sort((a, b) => a.name.localeCompare(b.name));
      case 'name-desc':
        return copy.sort((a, b) => b.name.localeCompare(a.name));
      case 'courses-desc':
        return copy.sort((a, b) => (b.courses.length || 0) - (a.courses.length || 0));
      case 'city-asc':
        return copy.sort((a, b) => (a.city || '').localeCompare(b.city || ''));
      default:
        return copy;
    }
  }

  // --- Render College Cards ---
  function renderCollegeCards(colleges) {
    // Empty state handling
    if (!colleges || colleges.length === 0) {
      resultsGrid.innerHTML = `
        <div class="empty-state-card" role="region" aria-label="No colleges found">
          <div class="empty-state-icon-wrap" aria-hidden="true">
            <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <circle cx="11" cy="11" r="8"/>
              <line x1="21" y1="21" x2="16.65" y2="16.65"/>
              <line x1="8" y1="11" x2="14" y2="11"/>
            </svg>
          </div>
          <h3 class="empty-state-title">No Colleges Found</h3>
          <p class="empty-state-desc">
            We couldn't find any college matching your current criteria. Try adjusting your search query or clearing specific filters.
          </p>
          <button type="button" class="btn-empty-reset" id="empty-clear-btn">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
              <polyline points="23 4 23 10 17 10"/>
              <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>
            </svg>
            Clear All Filters
          </button>
        </div>
      `;

      resultsCount.textContent = '0 colleges found';
      const emptyClearBtn = document.getElementById('empty-clear-btn');
      if (emptyClearBtn) {
        emptyClearBtn.addEventListener('click', clearAllFilters);
      }
      return;
    }

    // Results Counter announcement
    const countText = `${colleges.length} college${colleges.length === 1 ? '' : 's'} found`;
    resultsCount.textContent = countText;

    // Render Cards
    const cardsHtml = colleges.map((col, idx) => {
      const topCourses = col.courses.slice(0, 3);
      const remainingCount = col.courses.length - topCourses.length;

      const locationText = [col.city, col.state].filter(Boolean).join(', ') || 'India';
      const contactDisplay = col.contact_number
        ? `<div class="card-contact-row">
             <a href="tel:${escapeHtml(col.contact_number)}" class="contact-link" title="Call ${escapeHtml(col.name)}" onclick="event.stopPropagation()">
               <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/></svg>
               <span>${escapeHtml(col.contact_number)}</span>
             </a>
           </div>`
        : `<div class="card-contact-row" style="color: var(--text-dim);">No contact number listed</div>`;

      const tagsHtml = topCourses.map(course =>
        `<span class="course-chip">${escapeHtml(course)}</span>`
      ).join('');

      const moreHtml = remainingCount > 0
        ? `<span class="more-chip">+${remainingCount} more</span>`
        : '';

      return `
        <article
          class="college-card"
          tabindex="0"
          role="button"
          aria-haspopup="dialog"
          aria-label="${escapeHtml(col.name)} in ${escapeHtml(locationText)}. Click for full courses and details."
          data-id="${col.id}"
        >
          <div class="card-top-row">
            <span class="location-tag">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
              ${escapeHtml(locationText)}
            </span>
            <span class="courses-count-pill">${col.courses.length} courses</span>
          </div>

          <h3 class="card-college-name">${escapeHtml(col.name)}</h3>

          ${contactDisplay}

          <div class="card-courses-section">
            <div class="card-courses-heading">Offered Courses:</div>
            <div class="card-tags-list">
              ${tagsHtml}
              ${moreHtml}
            </div>
          </div>
        </article>
      `;
    }).join('');

    resultsGrid.innerHTML = cardsHtml;

    // Attach card click & keyboard listeners
    const cardElements = resultsGrid.querySelectorAll('.college-card');
    cardElements.forEach(card => {
      const id = parseInt(card.dataset.id, 10);
      const collegeData = allColleges.find(c => c.id === id);

      if (collegeData) {
        card.addEventListener('click', () => openDetailModal(collegeData));
        card.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            openDetailModal(collegeData);
          }
        });
      }
    });
  }

  // --- Update Active Badges Tray ---
  function updateActiveBadges() {
    const badges = [];

    if (filterState.query.trim()) {
      badges.push({
        type: 'query',
        label: `Search: "${filterState.query.trim()}"`
      });
    }
    if (filterState.state) {
      badges.push({
        type: 'state',
        label: `State: ${filterState.state}`
      });
    }
    if (filterState.city) {
      badges.push({
        type: 'city',
        label: `City: ${filterState.city}`
      });
    }
    if (filterState.course) {
      badges.push({
        type: 'course',
        label: `Course: ${filterState.course}`
      });
    }

    if (badges.length === 0) {
      activeTray.classList.add('hidden');
      activeBadgesList.innerHTML = '';
      return;
    }

    activeTray.classList.remove('hidden');
    activeBadgesList.innerHTML = badges.map(b => `
      <span class="filter-badge-chip">
        ${escapeHtml(b.label)}
        <button
          type="button"
          class="badge-dismiss-btn"
          data-type="${b.type}"
          aria-label="Remove filter: ${escapeHtml(b.label)}"
          title="Remove filter"
        >✕</button>
      </span>
    `).join('');

    // Attach dismiss clicks
    activeBadgesList.querySelectorAll('.badge-dismiss-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const type = e.target.dataset.type;
        if (type === 'query') {
          filterState.query = '';
          searchInput.value = '';
          searchClearBtn.classList.add('hidden');
        } else if (type === 'state') {
          filterState.state = '';
          stateSelect.value = '';
          repopulateCityDropdown();
        } else if (type === 'city') {
          filterState.city = '';
          citySelect.value = '';
        } else if (type === 'course') {
          filterState.course = '';
          courseSelect.value = '';
        }
        applyFilters();
      });
    });
  }

  // --- Reset All Filters ---
  function clearAllFilters() {
    filterState.query = '';
    filterState.state = '';
    filterState.city = '';
    filterState.course = '';

    searchInput.value = '';
    searchClearBtn.classList.add('hidden');
    stateSelect.value = '';
    courseSelect.value = '';
    repopulateCityDropdown();
    citySelect.value = '';

    applyFilters();
    searchInput.focus();
  }

  // --- Modal Dialog Handling ---
  let previouslyFocusedElement = null;

  function openDetailModal(college) {
    previouslyFocusedElement = document.activeElement;

    const locationFull = [college.city, college.state].filter(Boolean).join(', ') || 'India';
    const coursesListHtml = college.courses && college.courses.length > 0
      ? college.courses.map(c => `<span class="modal-course-pill">${escapeHtml(c)}</span>`).join('')
      : '<p style="color: var(--text-dim); font-size: 0.875rem;">No courses listed for this college.</p>';

    modalContentSlot.innerHTML = `
      <span class="modal-badge">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
        ${escapeHtml(college.state || 'India')}
      </span>
      <h2 id="modal-college-title" class="modal-title">${escapeHtml(college.name)}</h2>

      <div class="modal-info-tiles">
        <div class="modal-tile">
          <div class="tile-label">Location</div>
          <div class="tile-value">${escapeHtml(locationFull)}</div>
        </div>
        <div class="modal-tile">
          <div class="tile-label">Official Contact</div>
          <div class="tile-value">
            ${college.contact_number
              ? `<a href="tel:${escapeHtml(college.contact_number)}" class="contact-link" style="color: var(--accent-secondary);">
                   ${escapeHtml(college.contact_number)}
                 </a>`
              : '<span style="color: var(--text-dim);">Not Available</span>'}
          </div>
        </div>
      </div>

      <div class="modal-courses-section">
        <h3 class="modal-courses-header">Available Courses & Degrees (${college.courses.length})</h3>
        <div class="modal-courses-cloud">
          ${coursesListHtml}
        </div>
      </div>
    `;

    modalBackdrop.classList.add('is-open');
    modalBackdrop.setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';

    // Move focus to close button for screen reader convenience
    modalCloseBtn.focus();
  }

  function closeDetailModal() {
    modalBackdrop.classList.remove('is-open');
    modalBackdrop.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';

    if (previouslyFocusedElement) {
      previouslyFocusedElement.focus();
    }
  }

  // --- Event Listeners ---

  // Search Input with Debounce
  const handleSearchDebounced = debounce((value) => {
    filterState.query = value;
    applyFilters();
  }, 220);

  searchInput.addEventListener('input', (e) => {
    const val = e.target.value;
    if (val.trim()) {
      searchClearBtn.classList.remove('hidden');
    } else {
      searchClearBtn.classList.add('hidden');
    }
    handleSearchDebounced(val);
  });

  // Clear search input button
  searchClearBtn.addEventListener('click', () => {
    searchInput.value = '';
    searchClearBtn.classList.add('hidden');
    filterState.query = '';
    applyFilters();
    searchInput.focus();
  });

  // State Filter Change
  stateSelect.addEventListener('change', (e) => {
    filterState.state = e.target.value;
    // If state changes, adjust available cities
    filterState.city = '';
    repopulateCityDropdown();
    applyFilters();
  });

  // City Filter Change
  citySelect.addEventListener('change', (e) => {
    filterState.city = e.target.value;
    applyFilters();
  });

  // Course Filter Change
  courseSelect.addEventListener('change', (e) => {
    filterState.course = e.target.value;
    applyFilters();
  });

  // Sort Order Change
  sortSelect.addEventListener('change', (e) => {
    filterState.sortBy = e.target.value;
    applyFilters();
  });

  // Reset Button
  resetBtn.addEventListener('click', clearAllFilters);

  // Modal Close Events
  modalCloseBtn.addEventListener('click', closeDetailModal);
  modalBackdrop.addEventListener('click', (e) => {
    if (e.target === modalBackdrop) {
      closeDetailModal();
    }
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modalBackdrop.classList.contains('is-open')) {
      closeDetailModal();
    }
  });

  // Kickstart Application
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
  } else {
    initApp();
  }
})();