function toggleSidebar() {
  const sidebar = document.getElementById("sidebar");
  const overlay = document.querySelector(".sidebar-overlay");
  sidebar.classList.toggle("active");
  overlay.classList.toggle("active");
}

document.getElementById("hamburger")?.addEventListener("click", toggleSidebar);

// Slide animation for submenu
document.querySelectorAll(".has-submenu > a").forEach(link => {
  link.addEventListener("click", function (e) {
    e.preventDefault();
    const parent = this.parentElement;
    const submenu = parent.querySelector(".sub-menu");

    if (!submenu) return;

    const isExpanded = parent.classList.contains("expanded");

    submenu.style.overflow = "hidden";
    submenu.style.height = isExpanded ? submenu.scrollHeight + "px" : "0px";

    requestAnimationFrame(() => {
      submenu.style.transition = "height 0.3s ease";
      submenu.style.height = isExpanded ? "0px" : submenu.scrollHeight + "px";
    });

    submenu.addEventListener("transitionend", () => {
      submenu.style.height = isExpanded ? "0px" : "auto";
    }, { once: true });

    parent.classList.toggle("expanded");
  });
});

// Collapse sidebar on menu link click (mobile only)
document.querySelectorAll(".sidebar ul li a").forEach(link => {
  link.addEventListener("click", () => {
    const sidebar = document.getElementById("sidebar");
    const overlay = document.querySelector(".sidebar-overlay");
    if (window.innerWidth <= 992) {
      sidebar.classList.remove("active");
      overlay.classList.remove("active");
    }
  });

  // Animate icon
  link.addEventListener("mouseenter", () => {
    link.querySelector(".icon")?.classList.add("fa-beat");
  });
  link.addEventListener("mouseleave", () => {
    link.querySelector(".icon")?.classList.remove("fa-beat");
  });
});
