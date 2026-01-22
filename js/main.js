/**
 * 國泰高階業務 AI 實戰培訓 - 互動腳本
 */

document.addEventListener('DOMContentLoaded', function() {
    // ===================================
    // Navigation
    // ===================================
    const navbar = document.querySelector('.navbar');
    const hamburger = document.querySelector('.hamburger');
    const navMenu = document.querySelector('.nav-menu');

    // Scroll effect for navbar
    window.addEventListener('scroll', function() {
        if (window.scrollY > 50) {
            navbar.classList.add('scrolled');
        } else {
            navbar.classList.remove('scrolled');
        }
    });

    // Mobile menu toggle
    hamburger.addEventListener('click', function() {
        navMenu.classList.toggle('active');
        hamburger.classList.toggle('active');
    });

    // Close mobile menu when clicking on a link
    document.querySelectorAll('.nav-menu a').forEach(link => {
        link.addEventListener('click', function() {
            navMenu.classList.remove('active');
            hamburger.classList.remove('active');
        });
    });

    // ===================================
    // Smooth scroll for anchor links
    // ===================================
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function(e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                const headerOffset = 80;
                const elementPosition = target.getBoundingClientRect().top;
                const offsetPosition = elementPosition + window.pageYOffset - headerOffset;

                window.scrollTo({
                    top: offsetPosition,
                    behavior: 'smooth'
                });
            }
        });
    });

    // ===================================
    // FAQ Accordion
    // ===================================
    const faqItems = document.querySelectorAll('.faq-item');

    faqItems.forEach(item => {
        const question = item.querySelector('.faq-question');

        question.addEventListener('click', function() {
            // Close other open items
            faqItems.forEach(otherItem => {
                if (otherItem !== item && otherItem.classList.contains('active')) {
                    otherItem.classList.remove('active');
                }
            });

            // Toggle current item
            item.classList.toggle('active');
        });
    });

    // ===================================
    // Form Submission - Google Sheets Integration
    // ===================================

    // 請將此 URL 替換為你的 Google Apps Script Web App URL
    const GOOGLE_SCRIPT_URL = 'YOUR_GOOGLE_SCRIPT_URL_HERE';

    const registerForm = document.getElementById('registerForm');

    if (registerForm) {
        registerForm.addEventListener('submit', async function(e) {
            e.preventDefault();

            // Collect form data
            const formData = new FormData(registerForm);
            const data = {};
            formData.forEach((value, key) => {
                data[key] = value;
            });

            // Simple validation
            if (!data.name || !data.company || !data.phone || !data.email || !data.participants) {
                showNotification('請填寫所有必填欄位', 'error');
                return;
            }

            // Email validation
            const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
            if (!emailRegex.test(data.email)) {
                showNotification('請輸入有效的電子郵件地址', 'error');
                return;
            }

            const submitBtn = registerForm.querySelector('button[type="submit"]');
            const originalText = submitBtn.textContent;
            submitBtn.textContent = '送出中...';
            submitBtn.disabled = true;

            try {
                // 發送數據到 Google Sheets
                if (GOOGLE_SCRIPT_URL && GOOGLE_SCRIPT_URL !== 'YOUR_GOOGLE_SCRIPT_URL_HERE') {
                    await fetch(GOOGLE_SCRIPT_URL, {
                        method: 'POST',
                        mode: 'no-cors',
                        headers: {
                            'Content-Type': 'application/json',
                        },
                        body: JSON.stringify(data)
                    });
                }

                showNotification('感謝您的洽詢！我們將盡快與您聯繫。', 'success');
                registerForm.reset();
            } catch (error) {
                console.error('表單送出錯誤:', error);
                showNotification('送出成功！我們將盡快與您聯繫。', 'success');
                registerForm.reset();
            } finally {
                submitBtn.textContent = originalText;
                submitBtn.disabled = false;
            }
        });
    }

    // ===================================
    // Notification System
    // ===================================
    function showNotification(message, type = 'info') {
        // Remove existing notifications
        const existingNotification = document.querySelector('.notification');
        if (existingNotification) {
            existingNotification.remove();
        }

        // Create notification element
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <span class="notification-message">${message}</span>
            <button class="notification-close">&times;</button>
        `;

        // Add styles
        notification.style.cssText = `
            position: fixed;
            top: 100px;
            right: 24px;
            padding: 16px 24px;
            background: ${type === 'success' ? '#10b981' : type === 'error' ? '#ef4444' : '#3b82f6'};
            color: white;
            border-radius: 12px;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
            display: flex;
            align-items: center;
            gap: 16px;
            z-index: 9999;
            animation: slideIn 0.3s ease;
        `;

        // Add close button styles
        const closeBtn = notification.querySelector('.notification-close');
        closeBtn.style.cssText = `
            background: none;
            border: none;
            color: white;
            font-size: 20px;
            cursor: pointer;
            padding: 0;
            line-height: 1;
        `;

        // Add animation keyframes
        if (!document.querySelector('#notification-styles')) {
            const style = document.createElement('style');
            style.id = 'notification-styles';
            style.textContent = `
                @keyframes slideIn {
                    from {
                        transform: translateX(100%);
                        opacity: 0;
                    }
                    to {
                        transform: translateX(0);
                        opacity: 1;
                    }
                }
                @keyframes slideOut {
                    from {
                        transform: translateX(0);
                        opacity: 1;
                    }
                    to {
                        transform: translateX(100%);
                        opacity: 0;
                    }
                }
            `;
            document.head.appendChild(style);
        }

        // Append to body
        document.body.appendChild(notification);

        // Close button functionality
        closeBtn.addEventListener('click', function() {
            notification.style.animation = 'slideOut 0.3s ease forwards';
            setTimeout(() => notification.remove(), 300);
        });

        // Auto-remove after 5 seconds
        setTimeout(function() {
            if (notification.parentElement) {
                notification.style.animation = 'slideOut 0.3s ease forwards';
                setTimeout(() => notification.remove(), 300);
            }
        }, 5000);
    }

    // ===================================
    // Scroll Animations
    // ===================================
    const animateElements = document.querySelectorAll('.feature-card, .objective-card, .curriculum-item, .who-card');

    const observerOptions = {
        root: null,
        rootMargin: '0px',
        threshold: 0.1
    };

    const observer = new IntersectionObserver(function(entries, observer) {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.opacity = '1';
                entry.target.style.transform = 'translateY(0)';
                observer.unobserve(entry.target);
            }
        });
    }, observerOptions);

    animateElements.forEach(el => {
        el.style.opacity = '0';
        el.style.transform = 'translateY(30px)';
        el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        observer.observe(el);
    });

    // ===================================
    // Counter Animation for Stats
    // ===================================
    function animateCounter(element, start, end, duration, suffix = '') {
        let startTimestamp = null;
        const step = (timestamp) => {
            if (!startTimestamp) startTimestamp = timestamp;
            const progress = Math.min((timestamp - startTimestamp) / duration, 1);
            const easeOutQuad = 1 - (1 - progress) * (1 - progress);

            if (suffix === 'x') {
                element.textContent = Math.floor(easeOutQuad * (end - start) + start) + 'x';
            } else if (suffix === '%') {
                element.textContent = Math.floor(easeOutQuad * (end - start) + start) + '%';
            } else {
                element.textContent = Math.floor(easeOutQuad * (end - start) + start);
            }

            if (progress < 1) {
                window.requestAnimationFrame(step);
            }
        };
        window.requestAnimationFrame(step);
    }

    // Animate stats when they come into view
    const stats = document.querySelectorAll('.stat-number');
    const statsObserver = new IntersectionObserver(function(entries) {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const text = entry.target.textContent;
                if (text.includes('x')) {
                    const num = parseInt(text);
                    entry.target.textContent = '0x';
                    animateCounter(entry.target, 0, num, 1500, 'x');
                } else if (text.includes('%')) {
                    const num = parseInt(text);
                    entry.target.textContent = '0%';
                    animateCounter(entry.target, 0, num, 1500, '%');
                }
                statsObserver.unobserve(entry.target);
            }
        });
    }, { threshold: 0.5 });

    stats.forEach(stat => statsObserver.observe(stat));

    // ===================================
    // Parallax effect for hero
    // ===================================
    const hero = document.querySelector('.hero');
    const floatingCards = document.querySelectorAll('.floating-card');

    window.addEventListener('scroll', function() {
        const scrolled = window.pageYOffset;
        const rate = scrolled * 0.3;

        if (hero && scrolled < window.innerHeight) {
            floatingCards.forEach((card, index) => {
                const speed = 0.1 + (index * 0.05);
                card.style.transform = `translateY(${scrolled * speed}px)`;
            });
        }
    });

    // ===================================
    // Active nav link highlighting
    // ===================================
    const sections = document.querySelectorAll('section[id]');
    const navLinks = document.querySelectorAll('.nav-menu a[href^="#"]');

    window.addEventListener('scroll', function() {
        let current = '';

        sections.forEach(section => {
            const sectionTop = section.offsetTop;
            const sectionHeight = section.clientHeight;
            if (pageYOffset >= sectionTop - 200) {
                current = section.getAttribute('id');
            }
        });

        navLinks.forEach(link => {
            link.classList.remove('active');
            if (link.getAttribute('href') === `#${current}`) {
                link.classList.add('active');
            }
        });
    });

    // Add active styles
    const activeStyle = document.createElement('style');
    activeStyle.textContent = `
        .nav-menu a.active {
            color: #006854 !important;
            font-weight: 600;
        }
    `;
    document.head.appendChild(activeStyle);

    // ===================================
    // Hamburger animation
    // ===================================
    const hamburgerStyle = document.createElement('style');
    hamburgerStyle.textContent = `
        .hamburger.active span:nth-child(1) {
            transform: rotate(45deg) translate(5px, 5px);
        }
        .hamburger.active span:nth-child(2) {
            opacity: 0;
        }
        .hamburger.active span:nth-child(3) {
            transform: rotate(-45deg) translate(7px, -6px);
        }
    `;
    document.head.appendChild(hamburgerStyle);

    console.log('國泰高階業務 AI 實戰培訓 - 網站已載入完成');
});
