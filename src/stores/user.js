import { defineStore } from 'pinia';
// 定义用户 Store
export const useUserStore = defineStore('user', {
    // 状态
    state: () => ({
        userId: '',
        username: '',
        isLoggedIn: false,
    }),
    // Getters（类似 computed）
    getters: {
        fullName: (state) => `${state.username} (ID: ${state.userId})`,
        isAuthenticated: (state) => state.isLoggedIn,
    },
    // Actions（方法）
    actions: {
        login(userId, username) {
            this.userId = userId;
            this.username = username;
            this.isLoggedIn = true;
        },
        logout() {
            this.userId = '';
            this.username = '';
            this.isLoggedIn = false;
        },
    },
});
//# sourceMappingURL=user.js.map