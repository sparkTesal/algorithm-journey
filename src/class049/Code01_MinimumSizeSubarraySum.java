package class049;

// 累加和>=target的长度最小子数组
// 给定一个含有 n 个正整数的数组和一个正整数 target
// 找到累加和 >= target 的长度最小的子数组并返回其长度
// 如果不存在符合条件的子数组返回0
// 测试链接 : https://leetcode.cn/problems/minimum-size-subarray-sum/
public class Code01_MinimumSizeSubarraySum {

	public static int minSubArrayLen(int target, int[] nums) {
		int ans = Integer.MAX_VALUE;
		for (int l = 0, r = 0, sum = 0; r < nums.length; r++) {
			sum += nums[r];
			while (sum - nums[l] >= target) {
				sum -= nums[l++];
			}
			if (sum >= target) {
				ans = Math.min(ans, r - l + 1);
			}
		}
		return ans == Integer.MAX_VALUE ? 0 : ans;
	}

}
