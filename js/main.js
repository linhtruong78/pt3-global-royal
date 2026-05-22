/* ============================================================
   PT3 GLOBAL ROYAL — Main JS
   GSAP animations + global interactions
   ============================================================ */

document.addEventListener('DOMContentLoaded', () => {

  // ----------------------------------------------------------
  // 1. GSAP — Register ScrollTrigger
  // ----------------------------------------------------------
  gsap.registerPlugin(ScrollTrigger);

  // ----------------------------------------------------------
  // 2. Fade up on scroll (applied to .fade-up elements)
  // ----------------------------------------------------------
  gsap.utils.toArray('.fade-up').forEach(el => {
    gsap.to(el, {
      opacity: 1,
      y: 0,
      duration: 1.2,
      ease: 'power3.out',
      scrollTrigger: {
        trigger: el,
        start: 'top 88%',
        toggleActions: 'play none none none'
      }
    });
  });

  // ----------------------------------------------------------
  // 3. Fade in on scroll (applied to .fade-in elements)
  // ----------------------------------------------------------
  gsap.utils.toArray('.fade-in').forEach(el => {
    gsap.to(el, {
      opacity: 1,
      duration: 1.4,
      ease: 'power2.out',
      scrollTrigger: {
        trigger: el,
        start: 'top 88%',
        toggleActions: 'play none none none'
      }
    });
  });

  // ----------------------------------------------------------
  // 4. Divider line draw animation
  // ----------------------------------------------------------
  gsap.utils.toArray('.divider-gold').forEach(el => {
    gsap.fromTo(el,
      { scaleX: 0, transformOrigin: 'left center' },
      {
        scaleX: 1,
        duration: 1.2,
        ease: 'power3.out',
        scrollTrigger: {
          trigger: el,
          start: 'top 90%',
          toggleActions: 'play none none none'
        }
      }
    );
  });

  // ----------------------------------------------------------
  // 5. Navbar scroll behavior — adds .scrolled class
  // ----------------------------------------------------------
  const navbar = document.querySelector('.navbar');
  if (navbar) {
    ScrollTrigger.create({
      start: 'top -80',
      onUpdate: self => {
        navbar.classList.toggle('scrolled', self.scroll() > 80);
      }
    });
  }

  // ----------------------------------------------------------
  // 6. Mobile menu toggle
  // ----------------------------------------------------------
  const menuToggle = document.querySelector('.menu-toggle');
  const mobileMenu = document.querySelector('.mobile-menu');

  if (menuToggle && mobileMenu) {
    menuToggle.addEventListener('click', () => {
      const isOpen = mobileMenu.classList.toggle('open');
      menuToggle.setAttribute('aria-expanded', isOpen);
      document.body.style.overflow = isOpen ? 'hidden' : '';
    });
  }

});
