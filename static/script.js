// script.js

document.addEventListener('DOMContentLoaded', () => {
  // -------------------------
  // Sidebar toggle (mobile)
  // -------------------------
  const sidebarToggle = document.getElementById('sidebar-toggle');
  const sidebar = document.getElementById('sidebar');
  if(sidebarToggle){
    sidebarToggle.addEventListener('click', () => {
      sidebar.classList.toggle('-translate-x-full');
    });
  }

  // -------------------------
  // Submenu toggle
  // -------------------------
  document.querySelectorAll('button[data-target]').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = document.getElementById(btn.dataset.target);
      const plus = btn.querySelector('.plus-icon');
      const minus = btn.querySelector('.minus-icon');

      target.classList.toggle('open');
      plus.classList.toggle('hidden');
      minus.classList.toggle('hidden');
    });
  });

  // -------------------------
  // Profile dropdown
  // -------------------------
  const profileBtn = document.getElementById('profile-button');
  const profileMenu = document.getElementById('profile-menu');
  if(profileBtn){
    document.addEventListener("click", (e) => {
      if(profileBtn.contains(e.target)){
        profileMenu.classList.toggle("hidden");
      } else {
        profileMenu.classList.add("hidden");
      }
    });
  }

  // -------------------------
  // Navigation & content load
  // -------------------------
  function setActiveLink(link){
    document.querySelectorAll('a[data-link]').forEach(l => l.classList.remove('bg-gray-200'));
    document.querySelector(`[data-link="${link}"]`)?.classList.add('bg-gray-200');
  }

  function loadContent(hash){
    switch(hash){
      case '#students-view-all': fetchContent('/students').then(data => {
        document.getElementById('main-content').innerHTML = data;
        setActiveLink('students-view-all');
      }); break;
      case '#student-manage-login': fetchContent('/students-login').then(data => {
        document.getElementById('main-content').innerHTML = data;
        setActiveLink('student-manage-login');
      }); break;
      case '#teachers-view-all': fetchContent('/teachers').then(data => {
        document.getElementById('main-content').innerHTML = data;
        setActiveLink('teachers-view-all');
      }); break;
      case '#teacher-manage-login': fetchContent('/teachers-login').then(data => {
        document.getElementById('main-content').innerHTML = data;
        setActiveLink('teacher-manage-login');
      }); break;
      case '#attendance': fetchContent('/attendance').then(data => {
        document.getElementById('main-content').innerHTML = data;
        setActiveLink('attendance');
      }); break;
      case '#dashboard':
      default: fetchContent('/dashboard').then(data => {
        document.getElementById('main-content').innerHTML = data;
        setActiveLink('dashboard');
      });
    }
  }

  async function fetchContent(route){
    try{
      const res = await fetch(route);
      return await res.text();
    } catch(e){
      console.error(e);
      return "<p class='text-red-600'>Error loading content.</p>";
    }
  }

  // -------------------------
  // Initial load
  // -------------------------
  loadContent(window.location.hash || '#dashboard');

  window.addEventListener('hashchange', () => {
    loadContent(window.location.hash);
  });

  // -------------------------
  // Hover & animation effects
  // -------------------------
  document.querySelectorAll('nav a').forEach(link => {
    link.addEventListener('mouseenter', () => {
      link.classList.add('bg-gray-200', 'transition', 'duration-300');
    });
    link.addEventListener('mouseleave', () => {
      if(!link.classList.contains('active')){
        link.classList.remove('bg-gray-200');
      }
    });
  });
});
